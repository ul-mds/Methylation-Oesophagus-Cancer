from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from .io import ESCA_DIR, align_omics_to_samples, build_master_clinical, read_expression_zscores, read_methylation, read_mutations, mutation_matrix
from .interactions import screen_mutation_histology_interactions
from .mediation import MediationConfig, mediation_scan_by_gene
from .models import ModelConfig, build_feature_table, fit_cox


OUT_DIR = Path(__file__).resolve().parent / "out"


def _ensure_out():
    OUT_DIR.mkdir(parents=True, exist_ok=True)


def _pick_gene_probe_map(meth: pd.DataFrame) -> dict[str, str]:
    """
    We don't have an annotation table inside the cBioPortal package.
    As a practical proxy, use probe IDs that contain a gene-like token:
      - if probe name includes ';' or '_' we try the first token
    This is imperfect, but lets us run the pipeline end-to-end.
    """
    gene_to_probe: dict[str, str] = {}
    for probe in meth.index.astype(str):
        token = probe.split(";")[0].split("_")[0].split(":")[0]
        if token and token not in gene_to_probe:
            gene_to_probe[token] = probe
    return gene_to_probe


def main():
    _ensure_out()
    master = build_master_clinical()

    expr = read_expression_zscores(which="all_samples")
    meth = read_methylation()
    master, expr, meth = align_omics_to_samples(master, expr=expr, meth=meth)

    master.to_csv(OUT_DIR / "master_clinical_samples.tsv", sep="\t", index=False)

    gene_to_probe = _pick_gene_probe_map(meth)

    # --- Mediation scans (pragmatic survival mediation proxy) ---
    # Use genes present in expression AND in our proxy map.
    genes = [g for g in expr.index.astype(str).tolist() if g in gene_to_probe]
    # keep it tractable for first pass: top 1500 genes by variance across samples
    vv = expr.loc[genes, master["SAMPLE_ID"].tolist()].var(axis=1, skipna=True).sort_values(ascending=False)
    genes = vv.head(1500).index.tolist()

    for hist in ["EAC/ESCA", "ESCC"]:
        med = mediation_scan_by_gene(
            master,
            expr,
            meth,
            gene_to_probe,
            histology_label=hist,
            config=MediationConfig(endpoint="OS"),
        )
        med.to_csv(OUT_DIR / f"mediation_proxy_{hist.replace('/','_')}_OS.tsv", sep="\t", index=False)

    # Compare ESCA vs ESCC indirect effects for genes found in either
    m1 = pd.read_csv(OUT_DIR / "mediation_proxy_EAC_ESCA_OS.tsv", sep="\t")
    m2 = pd.read_csv(OUT_DIR / "mediation_proxy_ESCC_OS.tsv", sep="\t")
    if not m1.empty and not m2.empty:
        mm = m1.merge(m2, on="gene", how="outer", suffixes=("_ESCA", "_ESCC"))
        mm["delta_indirect_ab"] = mm["indirect_ab_ESCA"] - mm["indirect_ab_ESCC"]
        mm.to_csv(OUT_DIR / "mediation_proxy_compare_ESCA_vs_ESCC_OS.tsv", sep="\t", index=False)

    # --- Prognostic models + cross-apply (simple baseline) ---
    # Use top 200 mediation-candidate genes (union) as features to keep p<<n.
    top_genes = []
    if (OUT_DIR / "mediation_proxy_EAC_ESCA_OS.tsv").exists():
        top_genes += pd.read_csv(OUT_DIR / "mediation_proxy_EAC_ESCA_OS.tsv", sep="\t").head(120)["gene"].tolist()
    if (OUT_DIR / "mediation_proxy_ESCC_OS.tsv").exists():
        top_genes += pd.read_csv(OUT_DIR / "mediation_proxy_ESCC_OS.tsv", sep="\t").head(120)["gene"].tolist()
    top_genes = list(dict.fromkeys([g for g in top_genes if isinstance(g, str)]))[:200]

    if len(top_genes) >= 20:
        X = build_feature_table(master, expr, meth, gene_to_probe, genes=top_genes)
        df_model = master.set_index("SAMPLE_ID").join(X, how="inner")
        duration_col, event_col = "OS_MONTHS", "OS_EVENT"
        df_model = df_model.dropna(subset=[duration_col, event_col])
        df_model[duration_col] = df_model[duration_col].astype(float)
        df_model[event_col] = df_model[event_col].astype(int)

        # ESCA-only
        esca = df_model[df_model["HISTOLOGY"] == "EAC/ESCA"].copy()
        escc = df_model[df_model["HISTOLOGY"] == "ESCC"].copy()

        cfg = ModelConfig(endpoint="OS", penalizer=0.1, l1_ratio=0.5)

        # Fit models and export coefficients
        if esca[event_col].sum() >= cfg.min_events:
            cph_esca = fit_cox(esca[[duration_col, event_col] + list(X.columns)], duration_col, event_col, cfg)
            cph_esca.summary.to_csv(OUT_DIR / "cox_model_ESCA_only_OS.tsv", sep="\t")
            # cross-apply: score ESCC
            if len(escc) > 0:
                risk_escc = cph_esca.predict_partial_hazard(escc[list(X.columns)]).iloc[:, 0]
                risk_escc.to_csv(OUT_DIR / "risk_ESCA_model_applied_to_ESCC.tsv", sep="\t")

        if escc[event_col].sum() >= cfg.min_events:
            cph_escc = fit_cox(escc[[duration_col, event_col] + list(X.columns)], duration_col, event_col, cfg)
            cph_escc.summary.to_csv(OUT_DIR / "cox_model_ESCC_only_OS.tsv", sep="\t")
            # cross-apply: score ESCA
            if len(esca) > 0:
                risk_esca = cph_escc.predict_partial_hazard(esca[list(X.columns)]).iloc[:, 0]
                risk_esca.to_csv(OUT_DIR / "risk_ESCC_model_applied_to_ESCA.tsv", sep="\t")

        # pooled with stratified baseline
        pooled = df_model[[duration_col, event_col, "HISTOLOGY"] + list(X.columns)].copy()
        if pooled[event_col].sum() >= cfg.min_events:
            cph_pool = fit_cox(pooled, duration_col, event_col, cfg, strata=["HISTOLOGY"])
            cph_pool.summary.to_csv(OUT_DIR / "cox_model_pooled_strataHIST_OS.tsv", sep="\t")

    # --- Mutation × histology interactions ---
    muts = read_mutations()
    # choose top mutated genes within available samples
    samples = master["SAMPLE_ID"].astype(str).tolist()
    muts_sub = muts[muts["Tumor_Sample_Barcode"].isin(samples)]
    top_mut_genes = muts_sub["Hugo_Symbol"].value_counts().head(80).index.tolist()
    mut_mat = mutation_matrix(muts_sub, samples=samples, genes=top_mut_genes)
    inter = screen_mutation_histology_interactions(master, mut_mat, genes=top_mut_genes)
    inter.to_csv(OUT_DIR / "mutation_x_histology_interactions_OS.tsv", sep="\t", index=False)


if __name__ == "__main__":
    main()

