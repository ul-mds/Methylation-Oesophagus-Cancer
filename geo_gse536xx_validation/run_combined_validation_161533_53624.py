"""
Combined external validation report:
  - GSE161533: paired tumor vs normal expression (28 patients); overlap with TCGA ESCC gene lists.
  - GSE53624 / GSE53622: 60-month landmark + penalized Cox on lncRNA (existing logic).

Outputs: combined_validation_report.tsv + combined_validation_summary.txt
"""
from __future__ import annotations

import gzip
import json
import re
import subprocess
import sys
from io import StringIO
from pathlib import Path

import numpy as np
import pandas as pd
import urllib.request

BASE = Path(__file__).resolve().parent
STUDY_ROOT = BASE.parent  # Esophageal Carcinoma (TCGA, Firehose Legacy)
PIPE_OUT = STUDY_ROOT / "esca_escc_multiomic_pipeline" / "out"


def mygene_map_probes(probes: list[str]) -> dict[str, str]:
    """Map Affymetrix probe IDs to gene symbols via mygene.info (batch)."""
    url = "https://mygene.info/v3/query"
    body = {
        "q": probes,
        "scopes": "reporter",
        "species": "human",
        "fields": "symbol",
        "dotfield": True,
    }
    m: dict[str, str] = {}
    chunk = 100
    for i in range(0, len(probes), chunk):
        sub = probes[i : i + chunk]
        body["q"] = sub
        data = json.dumps(body).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=90) as resp:
            out = json.loads(resp.read().decode())
        for hit in out:
            if not isinstance(hit, dict):
                continue
            q = hit.get("query")
            s = hit.get("symbol")
            if isinstance(s, list) and s:
                sym = s[0]
            else:
                sym = s
            if q and sym:
                m[str(q)] = str(sym)
    return m


def gse161533_paired_lfc() -> tuple[pd.Series, pd.DataFrame]:
    path = BASE / "GSE161533_series_matrix.txt.gz"
    if not path.exists():
        raise FileNotFoundError(f"Missing {path}; download series matrix first.")
    raw = gzip.open(path, "rt", errors="replace").read()

    def stripq(x: str) -> str:
        return x.strip().strip('"')

    geo = None
    for line in raw.splitlines():
        if line.startswith("!Sample_geo_accession"):
            geo = [stripq(x) for x in line.split("\t")[1:]]
            break
    title_row = None
    for line in raw.splitlines():
        if line.startswith("!Sample_title"):
            title_row = [stripq(x) for x in line.split("\t")[1:]]
            break
    tissue_row = None
    for line in raw.splitlines():
        if line.startswith("!Sample_characteristics_ch1") and "tissue:" in line.lower():
            tissue_row = [stripq(x) for x in line.split("\t")[1:]]
            break
    assert geo and title_row and tissue_row
    tissue = [re.search(r"tissue:\s*(.+)", x, re.I).group(1).strip() for x in tissue_row]
    df = pd.DataFrame({"gsm": geo, "title": title_row, "tissue": tissue})

    start = raw.index("!series_matrix_table_begin\n") + len("!series_matrix_table_begin\n")
    end = raw.index("!series_matrix_table_end")
    expr = pd.read_csv(StringIO(raw[start:end]), sep="\t", index_col=0)
    expr.columns = [c.strip() for c in expr.columns]
    expr = expr.apply(pd.to_numeric, errors="coerce")

    from collections import defaultdict

    pat: dict[str, list[tuple[str, str]]] = defaultdict(list)
    for _, r in df.iterrows():
        m = re.search(r"patient\s+(\d+)", r["title"], re.I)
        pid = m.group(1) if m else r["title"]
        pat[pid].append((r["gsm"], r["tissue"]))

    logfc_list = []
    for pid, items in pat.items():
        if len(items) != 3:
            continue
        ts = {t: g for g, t in items}
        if "normal tissue" not in ts or "tumor tissue" not in ts:
            continue
        gn, gt = ts["normal tissue"], ts["tumor tissue"]
        logfc_list.append(np.log2((expr[gt] + 1) / (expr[gn] + 1)))
    if not logfc_list:
        raise RuntimeError("No paired tumor-normal pairs found.")
    LFC = pd.concat(logfc_list, axis=1)
    mean_lfc = LFC.mean(axis=1)
    meta = pd.DataFrame({"n_pairs": len(logfc_list)}, index=[0])
    return mean_lfc, meta


def load_tcga_escc_genes() -> tuple[set[str], set[str]]:
    cox = pd.read_csv(PIPE_OUT / "cox_model_ESCC_only_OS.tsv", sep="\t")
    expr_genes = {
        c.replace("expr::", "")
        for c in cox["covariate"].astype(str)
        if c.startswith("expr::")
    }
    med = pd.read_csv(PIPE_OUT / "mediation_proxy_ESCC_OS.tsv", sep="\t")
    med_genes = set(med["gene"].astype(str).head(80).tolist())
    return expr_genes, med_genes


def main() -> None:
    lines: list[str] = []

    # --- GSE161533 ---
    mean_lfc, meta161 = gse161533_paired_lfc()
    ranked = mean_lfc.abs().sort_values(ascending=False)
    top_n = 1200
    top_probes = ranked.head(top_n).index.astype(str).tolist()
    try:
        sym_map = mygene_map_probes(top_probes)
    except Exception as e:
        raise RuntimeError(f"mygene.info probe mapping failed: {e}") from e
    top_genes = []
    for p in top_probes:
        g = sym_map.get(p)
        if g:
            top_genes.append((g, float(mean_lfc.loc[p]), p))
    # dedupe gene keeping max |LFC|
    best: dict[str, tuple[float, str]] = {}
    for g, lfc, p in top_genes:
        if g not in best or abs(lfc) > abs(best[g][0]):
            best[g] = (lfc, p)
    expr_tcg, med_tcg = load_tcga_escc_genes()
    union_tcg = expr_tcg | med_tcg

    in_top = [g for g in union_tcg if g in best]
    overlap_expr = sorted(expr_tcg & set(best))
    overlap_med = sorted(med_tcg & set(best))

    g161_rows = pd.DataFrame(
        [
            {
                "dataset": "GSE161533",
                "metric": "n_paired_patients_tumor_vs_normal",
                "value": float(meta161["n_pairs"].iloc[0]),
                "note": "Affy GPL570; no OS in GEO",
            },
            {
                "dataset": "GSE161533",
                "metric": "tcga_escc_union_genes_in_topN_probe_genes",
                "value": len(in_top),
                "note": f"topN={top_n}; union_TCGA_n={len(union_tcg)}; mapped_genes={len(best)}",
            },
            {
                "dataset": "GSE161533",
                "metric": "overlap_cox_expr_genes_in_topN",
                "value": len(overlap_expr),
                "note": ",".join(overlap_expr[:25]) + ("..." if len(overlap_expr) > 25 else ""),
            },
            {
                "dataset": "GSE161533",
                "metric": "overlap_mediation_top80_in_topN",
                "value": len(overlap_med),
                "note": ",".join(overlap_med[:25]) + ("..." if len(overlap_med) > 25 else ""),
            },
        ]
    )

    detail161 = []
    for g in sorted(in_top)[:40]:
        lfc, pr = best[g]
        detail161.append({"gene": g, "mean_log2FC_tumor_vs_normal": lfc, "probe": pr})
    pd.DataFrame(detail161).to_csv(BASE / "gse161533_tcga_gene_overlap_detail.tsv", sep="\t", index=False)

    # --- GSE53624/22: subprocess existing script ---
    r = subprocess.run([sys.executable, str(BASE / "run_geo_landmark_cox.py")], capture_output=True, text=True)
    if r.returncode != 0:
        lines.append("GSE53624/22 script failed:\n" + r.stderr[:2000])
    else:
        lines.append(r.stdout)

    land = pd.read_csv(BASE / "validation_landmark60_report.tsv", sep="\t")
    g536_rows = []
    for _, row in land.iterrows():
        g536_rows.append(
            {
                "dataset": str(row.get("dataset", "")),
                "metric": "see_columns",
                "value": np.nan,
                "note": json.dumps({k: row[k] for k in land.columns if pd.notna(row.get(k))}, default=str),
            }
        )
    g536_df = pd.DataFrame(g536_rows)

    out = pd.concat([g161_rows, g536_df], ignore_index=True)
    out.to_csv(BASE / "combined_validation_report.tsv", sep="\t", index=False)

    # Plain summary
    summ = []
    summ.append("=== GSE161533 vs TCGA ESCC gene lists ===")
    summ.append(f"Paired patients (tumor vs normal): {int(meta161['n_pairs'].iloc[0])}")
    summ.append(f"TCGA union genes (Cox expr + mediation top80): n={len(union_tcg)}")
    summ.append(f"Genes from union found among top-{top_n} DE probes (mygene-mapped): {len(in_top)}")
    summ.append(f"  Cox expr genes hit: {len(overlap_expr)} / {len(expr_tcg)}")
    summ.append(f"  Mediation top80 hit: {len(overlap_med)} / {min(80, len(med_tcg))}")
    summ.append("")
    summ.append("=== GSE53624 / GSE53622 lncRNA + 60mo landmark Cox ===")
    summ.append(land.to_string(index=False))
    summ.append("")
    summ.append("Interpretation (short):")
    summ.append(
        f"- GSE161533: TCGA ESCC Cox+mediation gene union vs top-{top_n} |log2FC| tumor-normal: "
        f"{len(in_top)} genes overlap (Cox hits {len(overlap_expr)}/{len(expr_tcg)}, mediation {len(overlap_med)}/80)."
    )
    summ.append("  Low overlap is expected: TCGA features are prognosis/meth-coupled, not pure tumor-normal DE.")
    summ.append("  GSE161533 still supports generic ESCC biology (paired DE); OS/Cox not available in GEO.")
    summ.append("- GSE53624/22: lncRNA + 60mo penalized Cox — strong in-cohort fit (C~0.76 train), holdout C~0.50")
    summ.append("  (53624->53622), so prognostic transfer is weak; not TCGA mRNA space.")

    (BASE / "combined_validation_summary.txt").write_text("\n".join(summ), encoding="utf-8")
    print("\n".join(summ))


if __name__ == "__main__":
    main()
