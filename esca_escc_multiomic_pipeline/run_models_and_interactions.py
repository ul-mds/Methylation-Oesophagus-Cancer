from __future__ import annotations

from pathlib import Path

import pandas as pd
from lifelines.utils import concordance_index

from .io import align_omics_to_samples, build_master_clinical, read_expression_zscores, read_methylation, read_mutations, mutation_matrix
from .interactions import screen_mutation_histology_interactions
from .models import ModelConfig, build_feature_table, fit_cox
from .survival_60mo import truncate_os_at_months


OUT_DIR = Path(__file__).resolve().parent / "out"

# Landmark OS for survival models / C-index / interaction Cox (months)
OS_LANDMARK_MONTHS = 60.0


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    master = build_master_clinical()
    expr = read_expression_zscores(which="all_samples")
    meth = read_methylation()
    master, expr, meth = align_omics_to_samples(master, expr=expr, meth=meth)

    # Features: union of top mediation hits (from already-produced files)
    top_genes = []
    p1 = OUT_DIR / "mediation_proxy_EAC_ESCA_OS.tsv"
    p2 = OUT_DIR / "mediation_proxy_ESCC_OS.tsv"
    if p1.exists():
        top_genes += pd.read_csv(p1, sep="\t").head(30)["gene"].tolist()
    if p2.exists():
        top_genes += pd.read_csv(p2, sep="\t").head(30)["gene"].tolist()
    top_genes = list(dict.fromkeys([g for g in top_genes if isinstance(g, str)]))[:60]

    if len(top_genes) >= 20:
        gene_to_probe = {g: g for g in top_genes if g in meth.index}
        # Fallback if methylation index isn't gene-like for most; then skip meth features.
        if len(gene_to_probe) < 10:
            gene_to_probe = {}

        if gene_to_probe:
            X = build_feature_table(master, expr, meth, gene_to_probe, genes=top_genes)
        else:
            samples = master["SAMPLE_ID"].astype(str).tolist()
            X = expr.loc[[g for g in top_genes if g in expr.index], samples].T
            X.columns = [f"expr::{c}" for c in X.columns]

        df_model = master.set_index("SAMPLE_ID").join(X, how="inner")
        duration_col, event_col = "OS_MONTHS", "OS_EVENT"
        df_model = df_model.dropna(subset=[duration_col, event_col])
        df_model[duration_col] = df_model[duration_col].astype(float)
        df_model[event_col] = df_model[event_col].astype(int)
        # 60-month landmark OS for Cox / C-index
        t60, e60 = truncate_os_at_months(df_model[duration_col], df_model[event_col], months=OS_LANDMARK_MONTHS)
        df_model["OS_MONTHS_60"] = t60.to_numpy()
        df_model["OS_EVENT_60"] = e60.to_numpy()
        duration_col, event_col = "OS_MONTHS_60", "OS_EVENT_60"
        # Lifelines can't handle NaNs in covariates; median-impute per feature.
        for c in X.columns:
            df_model[c] = pd.to_numeric(df_model[c], errors="coerce")
        df_model[list(X.columns)] = df_model[list(X.columns)].apply(lambda s: s.fillna(s.median()), axis=0)
        df_model = df_model.dropna(subset=list(X.columns))

        esca = df_model[df_model["HISTOLOGY"] == "EAC/ESCA"].copy()
        escc = df_model[df_model["HISTOLOGY"] == "ESCC"].copy()

        cfg = ModelConfig(endpoint="OS", penalizer=0.1, l1_ratio=0.5)
        metrics = []

        if esca[event_col].sum() >= cfg.min_events:
            cph_esca = fit_cox(esca[[duration_col, event_col] + list(X.columns)], duration_col, event_col, cfg)
            cph_esca.summary.to_csv(OUT_DIR / "cox_model_ESCA_only_OS.tsv", sep="\t")
            if len(escc) > 0:
                ph = cph_esca.predict_partial_hazard(escc[list(X.columns)])
                risk_escc = ph.iloc[:, 0] if hasattr(ph, "iloc") and getattr(ph, "ndim", 1) == 2 else ph
                risk_escc.to_csv(OUT_DIR / "risk_ESCA_model_applied_to_ESCC.tsv", sep="\t")
                cidx = concordance_index(escc[duration_col], -risk_escc, escc[event_col])
                metrics.append({"model": "ESCA_only", "test": "ESCC", "c_index": float(cidx), "n": int(len(escc))})
            # internal
            ph_tr = cph_esca.predict_partial_hazard(esca[list(X.columns)])
            risk_tr = ph_tr.iloc[:, 0] if hasattr(ph_tr, "iloc") and getattr(ph_tr, "ndim", 1) == 2 else ph_tr
            cidx_tr = concordance_index(esca[duration_col], -risk_tr, esca[event_col])
            metrics.append({"model": "ESCA_only", "test": "ESCA", "c_index": float(cidx_tr), "n": int(len(esca))})

        if escc[event_col].sum() >= cfg.min_events:
            cph_escc = fit_cox(escc[[duration_col, event_col] + list(X.columns)], duration_col, event_col, cfg)
            cph_escc.summary.to_csv(OUT_DIR / "cox_model_ESCC_only_OS.tsv", sep="\t")
            if len(esca) > 0:
                ph = cph_escc.predict_partial_hazard(esca[list(X.columns)])
                risk_esca = ph.iloc[:, 0] if hasattr(ph, "iloc") and getattr(ph, "ndim", 1) == 2 else ph
                risk_esca.to_csv(OUT_DIR / "risk_ESCC_model_applied_to_ESCA.tsv", sep="\t")
                cidx = concordance_index(esca[duration_col], -risk_esca, esca[event_col])
                metrics.append({"model": "ESCC_only", "test": "ESCA", "c_index": float(cidx), "n": int(len(esca))})
            ph_tr = cph_escc.predict_partial_hazard(escc[list(X.columns)])
            risk_tr = ph_tr.iloc[:, 0] if hasattr(ph_tr, "iloc") and getattr(ph_tr, "ndim", 1) == 2 else ph_tr
            cidx_tr = concordance_index(escc[duration_col], -risk_tr, escc[event_col])
            metrics.append({"model": "ESCC_only", "test": "ESCC", "c_index": float(cidx_tr), "n": int(len(escc))})

        pooled = df_model[[duration_col, event_col, "HISTOLOGY"] + list(X.columns)].copy()
        if pooled[event_col].sum() >= cfg.min_events:
            cph_pool = fit_cox(pooled, duration_col, event_col, cfg, strata=["HISTOLOGY"])
            cph_pool.summary.to_csv(OUT_DIR / "cox_model_pooled_strataHIST_OS.tsv", sep="\t")
            ph = cph_pool.predict_partial_hazard(pooled[list(X.columns)])
            risk = ph.iloc[:, 0] if hasattr(ph, "iloc") and getattr(ph, "ndim", 1) == 2 else ph
            cidx = concordance_index(pooled[duration_col], -risk, pooled[event_col])
            metrics.append({"model": "pooled_strataHIST", "test": "pooled", "c_index": float(cidx), "n": int(len(pooled))})

        if metrics:
            mdf = pd.DataFrame(metrics)
            mdf["endpoint"] = f"OS_{int(OS_LANDMARK_MONTHS)}mo_landmark"
            mdf.to_csv(OUT_DIR / "model_performance_cindex.tsv", sep="\t", index=False)

    # Mutation × histology interactions (same landmark OS)
    muts = read_mutations()
    samples = master["SAMPLE_ID"].astype(str).tolist()
    muts_sub = muts[muts["Tumor_Sample_Barcode"].isin(samples)]
    top_mut_genes = muts_sub["Hugo_Symbol"].value_counts().head(80).index.tolist()
    mut_mat = mutation_matrix(muts_sub, samples=samples, genes=top_mut_genes)
    master_i = master.copy()
    t60m, e60m = truncate_os_at_months(master_i["OS_MONTHS"], master_i["OS_EVENT"], months=OS_LANDMARK_MONTHS)
    master_i["OS_MONTHS_60"] = t60m.to_numpy()
    master_i["OS_EVENT_60"] = e60m.to_numpy()
    inter = screen_mutation_histology_interactions(
        master_i, mut_mat, genes=top_mut_genes, time_col="OS_MONTHS_60", event_col="OS_EVENT_60"
    )
    inter.to_csv(OUT_DIR / "mutation_x_histology_interactions_OS.tsv", sep="\t", index=False)


if __name__ == "__main__":
    main()

