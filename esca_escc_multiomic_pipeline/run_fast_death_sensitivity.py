"""
Fast-death death-tail sensitivity analysis (10%, 20%, 30%).

Outputs under result/:
  - fast_death_sensitivity_summary.tsv   (patient-level + probe overlap)
  - fast_death_sensitivity_probes.tsv    (per-threshold top probes)
  - fast_death_sensitivity_key_loci.tsv  (leading loci rank stability)
"""
from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

from .io import ESCA_DIR, align_omics_to_samples, build_master_clinical, read_methylation
from .survival_60mo import truncate_os_at_months

THRESHOLDS = (0.10, 0.20, 0.30)
MIN_FAST_N = 10
TOP_PROBE_N = 200
KEY_GENES = (
    "SHISA3",
    "PRKG1",
    "C7orf13",
    "SEMA3E",
    "SIGLEC14",
    "EPHB1",
    "ZNF214",
    "RAB11FIP3",
)
OUT_DIR = ESCA_DIR / "result"


def _gene_token(probe: str) -> str:
    return str(probe).split(";")[0].split("_")[0].split(":")[0]


def _fast_death_ids(decedents: pd.DataFrame, pct: float, min_n: int = MIN_FAST_N) -> set[str]:
    n = max(min_n, int(math.ceil(len(decedents) * pct)))
    n = min(n, len(decedents))
    cutoff = decedents["OS60"].nsmallest(n).max()
    return set(decedents.loc[decedents["OS60"] <= cutoff, "SAMPLE_ID"].astype(str))


def _differential_methylation(
    meth: pd.DataFrame,
    sample_ids: list[str],
    fast_ids: set[str],
) -> pd.DataFrame:
    fast_cols = [s for s in sample_ids if s in fast_ids]
    other_cols = [s for s in sample_ids if s not in fast_ids]
    fast_mat = meth.loc[:, fast_cols].astype(float)
    other_mat = meth.loc[:, other_cols].astype(float)

    mean_fast = fast_mat.mean(axis=1)
    mean_other = other_mat.mean(axis=1)
    delta = mean_fast - mean_other

    pvals = []
    for probe in meth.index:
        a = fast_mat.loc[probe].dropna()
        b = other_mat.loc[probe].dropna()
        if len(a) >= 2 and len(b) >= 2:
            _, p = stats.ttest_ind(a, b, equal_var=False, nan_policy="omit")
        else:
            p = np.nan
        pvals.append(p)

    p = pd.Series(pvals, index=meth.index, dtype=float)
    q = pd.Series(stats.false_discovery_control(p.fillna(1.0).to_numpy(), method="bh"), index=meth.index)
    q[p.isna()] = np.nan

    return pd.DataFrame(
        {
            "probe": meth.index.astype(str),
            "gene_token": [_gene_token(x) for x in meth.index.astype(str)],
            "mean_beta_fast": mean_fast.to_numpy(),
            "mean_beta_other": mean_other.to_numpy(),
            "delta_fast_minus_other": delta.to_numpy(),
            "p_diff": p.to_numpy(),
            "q_diff": q.to_numpy(),
            "n_fast": len(fast_cols),
            "n_other": len(other_cols),
        }
    )


def _jaccard(a: set[str], b: set[str]) -> float:
    if not a and not b:
        return 1.0
    return len(a & b) / len(a | b)


def _top_probe_set(df: pd.DataFrame, n: int = TOP_PROBE_N) -> set[str]:
    sig = df.dropna(subset=["q_diff"]).sort_values("q_diff")
    return set(sig.head(n)["probe"].astype(str))


def _rank_of_gene(df: pd.DataFrame, gene: str) -> int | None:
    sub = df[df["gene_token"] == gene].dropna(subset=["q_diff"]).sort_values("q_diff")
    if sub.empty:
        return None
    ranks = df.dropna(subset=["q_diff"]).sort_values("q_diff").reset_index(drop=True)
    hit = ranks.index[ranks["gene_token"] == gene]
    return int(hit[0]) + 1 if len(hit) else None


def run_histology(
    hist: str,
    master: pd.DataFrame,
    meth: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sub = master[master["HISTOLOGY"] == hist].copy()
    dec = sub[sub["E60"] == 1].copy()
    sample_ids = sub["SAMPLE_ID"].astype(str).tolist()

    summary_rows: list[dict] = []
    probe_rows: list[pd.DataFrame] = []
    key_rows: list[dict] = []
    top_sets: dict[float, set[str]] = {}
    fast_sets: dict[float, set[str]] = {}
    diff_by_pct: dict[float, pd.DataFrame] = {}

    for pct in THRESHOLDS:
        fast = _fast_death_ids(dec, pct)
        fast_sets[pct] = fast
        diff = _differential_methylation(meth, sample_ids, fast)
        diff_by_pct[pct] = diff
        top_sets[pct] = _top_probe_set(diff)

        summary_rows.append(
            {
                "histology": hist,
                "death_tail_pct": int(pct * 100),
                "n_cohort": len(sub),
                "n_os60_deaths": len(dec),
                "n_fast_death": len(fast),
                "max_os_months_in_fast_group": float(dec.loc[dec["SAMPLE_ID"].isin(fast), "OS60"].max()),
                "n_probes_q_lt_005": int((diff["q_diff"] < 0.05).sum()),
                "n_top200_probes": len(top_sets[pct]),
            }
        )

        d = diff.copy()
        d["histology"] = hist
        d["death_tail_pct"] = int(pct * 100)
        probe_rows.append(d)

        for gene in KEY_GENES:
            gdf = diff[diff["gene_token"] == gene]
            if gdf.empty:
                continue
            row = gdf.sort_values("q_diff").iloc[0]
            key_rows.append(
                {
                    "histology": hist,
                    "death_tail_pct": int(pct * 100),
                    "gene": gene,
                    "probe": row["probe"],
                    "delta_beta": row["delta_fast_minus_other"],
                    "q_diff": row["q_diff"],
                    "genome_wide_q_rank": _rank_of_gene(diff, gene),
                }
            )

    pairs = [(10, 20), (10, 30), (20, 30)]
    for a, b in pairs:
        pa, pb = a / 100.0, b / 100.0
        summary_rows.append(
            {
                "histology": hist,
                "death_tail_pct": f"{a}_vs_{b}",
                "n_cohort": np.nan,
                "n_os60_deaths": len(dec),
                "n_fast_death": np.nan,
                "max_os_months_in_fast_group": np.nan,
                "n_probes_q_lt_005": np.nan,
                "n_top200_probes": np.nan,
                "patient_membership_jaccard": _jaccard(fast_sets[pa], fast_sets[pb]),
                "top200_probe_jaccard": _jaccard(top_sets[pa], top_sets[pb]),
            }
        )

    return (
        pd.DataFrame(summary_rows),
        pd.concat(probe_rows, ignore_index=True),
        pd.DataFrame(key_rows),
    )


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    master = build_master_clinical()
    meth = read_methylation()
    master, _, meth = align_omics_to_samples(master, meth=meth)
    master["OS60"], master["E60"] = truncate_os_at_months(
        master["OS_MONTHS"], master["OS_EVENT"], 60.0
    )

    summaries = []
    probes = []
    keys = []
    for hist in ["EAC/ESCA", "ESCC"]:
        s, p, k = run_histology(hist, master, meth)
        summaries.append(s)
        probes.append(p)
        keys.append(k)

    summary = pd.concat(summaries, ignore_index=True)
    probe_out = pd.concat(probes, ignore_index=True)
    key_out = pd.concat(keys, ignore_index=True)

    summary.to_csv(OUT_DIR / "fast_death_sensitivity_summary.tsv", sep="\t", index=False)
    probe_out.to_csv(OUT_DIR / "fast_death_sensitivity_probes.tsv", sep="\t", index=False)
    key_out.to_csv(OUT_DIR / "fast_death_sensitivity_key_loci.tsv", sep="\t", index=False)

    print("Wrote:")
    print(" ", OUT_DIR / "fast_death_sensitivity_summary.tsv")
    print(" ", OUT_DIR / "fast_death_sensitivity_probes.tsv")
    print(" ", OUT_DIR / "fast_death_sensitivity_key_loci.tsv")
    print("\nSummary preview:")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
