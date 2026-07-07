"""
GEO GSE53624 / GSE53622 / GSE53625-style validation (ESCC lncRNA cohorts).

This does NOT replicate TCGA multi-omic mediation/mutation analyses (no HM450, no MAF,
and GPL18109 is lncRNA-centric — not the TCGA protein-coding z-score space).

It applies the same *survival framing* as the manuscript pipeline:
  - 60-month landmark OS truncation
  - penalized Cox (elastic-net style via lifelines CoxPHFitter penalizer/l1_ratio)

Train: GSE53624 supplementary lncRNA matrix + clinical.
Test:  GSE53622 (features fixed from training variance ranking).
GSE53625: conceptual super-series of 53624+53622 — no separate clinical matrix here;
          reported as combined N for context.
"""
from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd
from lifelines import CoxPHFitter
from lifelines.utils import concordance_index

BASE = Path(__file__).resolve().parent
PENALIZER = 0.1
L1_RATIO = 0.5
LANDMARK = 60.0
TOP_K = 60
MIN_EVENTS = 15


def landmark_os(months: pd.Series, death_event: pd.Series, cap: float = LANDMARK) -> tuple[pd.Series, pd.Series]:
    t = pd.to_numeric(months, errors="coerce").astype(float)
    e = death_event.astype(int)
    t_trunc = np.minimum(t.to_numpy(), cap)
    # event only if death observed by cap
    e_trunc = np.where((e.to_numpy() == 1) & (t.to_numpy() <= cap), 1, 0)
    return pd.Series(t_trunc, index=months.index), pd.Series(e_trunc, index=months.index)


def load_expr_wide(xlsx: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name=0)
    id_col = df.columns[0]
    df = df.set_index(id_col)
    df.index = df.index.astype(str)
    df.columns = [str(c) for c in df.columns]
    return df


def load_clinical(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    death = df["Death at FU"].astype(str).str.lower().str.strip()
    event = death.eq("yes").astype(int)
    months = pd.to_numeric(df["Survival time(months)"], errors="coerce")
    out = pd.DataFrame(
        {
            "patient_col": df["Patient ID"].astype(str),
            "OS_MONTHS": months,
            "OS_EVENT": event,
        }
    )
    return out.dropna(subset=["OS_MONTHS"])


def build_xy(expr: pd.DataFrame, clin: pd.DataFrame) -> tuple[pd.DataFrame, pd.Series, pd.Series]:
    """expr: features x patients (columns ec*). Returns X (n x p), T, E."""
    cols = [c for c in expr.columns if c in set(clin["patient_col"])]
    miss = set(clin["patient_col"]) - set(cols)
    c2 = clin.set_index("patient_col").loc[cols]
    T, E = landmark_os(c2["OS_MONTHS"], c2["OS_EVENT"])
    X = expr[cols].T.astype(float)
    X.index = cols
    return X, T, E


def pick_top_k_features(expr_train: pd.DataFrame, k: int) -> list[str]:
    v = expr_train.var(axis=1, skipna=True).sort_values(ascending=False)
    return v.head(k).index.astype(str).tolist()


def main() -> None:
    expr24 = load_expr_wide(BASE / "GSE53624_lnc.xlsx")
    expr22 = load_expr_wide(BASE / "GSE53622_lnc.xlsx")
    clin24 = load_clinical(BASE / "GSE53624_clinical.xlsx")
    clin22 = load_clinical(BASE / "GSE53622_clinical.xlsx")

    X_tr, T_tr, E_tr = build_xy(expr24, clin24)
    cols_tr = [c for c in expr24.columns if c in set(clin24["patient_col"])]
    feats = pick_top_k_features(expr24[cols_tr], TOP_K)
    X_tr = X_tr[feats]

    if int(E_tr.sum()) < MIN_EVENTS:
        raise SystemExit(f"Training events at {LANDMARK}mo too few: {int(E_tr.sum())}")

    train_df = X_tr.copy()
    train_df["T"] = T_tr
    train_df["E"] = E_tr

    cph = CoxPHFitter(penalizer=PENALIZER, l1_ratio=L1_RATIO)
    cph.fit(train_df, duration_col="T", event_col="E")
    c_train = float(cph.concordance_index_)

    X_te, T_te, E_te = build_xy(expr22, clin22)
    common = [f for f in feats if f in expr22.index]
    if len(common) < 10:
        raise SystemExit("Too few overlapping feature IDs between cohorts.")
    X_te = X_te[common]
    # align columns to training model: lifelines needs same names; refit quick fix — use train subset
    train_sub = [f for f in feats if f in common]
    train_df2 = X_tr[train_sub].copy()
    train_df2["T"] = T_tr
    train_df2["E"] = E_tr
    cph2 = CoxPHFitter(penalizer=PENALIZER, l1_ratio=L1_RATIO)
    cph2.fit(train_df2, duration_col="T", event_col="E")

    X_te2 = X_te[train_sub]
    ph = cph2.predict_partial_hazard(X_te2)
    risk = ph.iloc[:, 0] if isinstance(ph, pd.DataFrame) else ph
    c_test = concordance_index(T_te.loc[X_te2.index], -risk, E_te.loc[X_te2.index])

    # Combined cohort stats (GSE53625 conceptual)
    n_combined = len(clin24) + len(clin22)
    ev24 = int(E_tr.sum())
    ev22 = int(E_te.sum())

    rep_path = BASE / "validation_landmark60_report.tsv"
    pd.DataFrame(
        [
            {
                "dataset": "GSE53624_train",
                "n_patients": len(X_tr),
                "events_60mo": ev24,
                "c_index_in_sample": c_train,
                "note": "Top-variance lncRNA features; not TCGA gene space",
            },
            {
                "dataset": "GSE53622_test",
                "n_patients": len(X_te2),
                "events_60mo": ev22,
                "c_index_holdout": float(c_test),
                "n_features_used": len(train_sub),
                "note": "Model trained GSE53624; features from train only",
            },
            {
                "dataset": "GSE53625_super_series",
                "n_patients": n_combined,
                "events_60mo_approx": ev24 + ev22,
                "note": "Union of 53624+53622; same as analyzing subseries separately",
            },
        ]
    ).to_csv(rep_path, sep="\t", index=False)

    print("Wrote", rep_path)
    print(f"GSE53624: n={len(X_tr)} events@60mo={ev24} C-index fit={c_train:.3f}")
    print(f"GSE53622: n={len(X_te2)} events@60mo={ev22} C-index holdout={c_test:.3f} (p={len(train_sub)})")

    # In-sample GSE53622 (same statistical shell; features from 53622 only)
    cols_22 = [c for c in expr22.columns if c in set(clin22["patient_col"])]
    feats22 = pick_top_k_features(expr22[cols_22], TOP_K)
    X_22, T_22, E_22 = build_xy(expr22, clin22)
    X_22 = X_22[feats22]
    c_22_in = np.nan
    if int(E_22.sum()) >= MIN_EVENTS:
        df22 = X_22.copy()
        df22["T"] = T_22
        df22["E"] = E_22
        cph22 = CoxPHFitter(penalizer=PENALIZER, l1_ratio=L1_RATIO)
        cph22.fit(df22, duration_col="T", event_col="E")
        c_22_in = float(cph22.concordance_index_)
        print(f"GSE53622 in-sample only: n={len(X_22)} C-index={c_22_in:.3f}")

    append_rows = [
        {
            "dataset": "GSE53622_in_sample",
            "n_patients": len(X_22),
            "events_60mo": int(E_22.sum()),
            "c_index_in_sample": c_22_in,
            "note": "Top-variance lncRNA fit only on GSE53622 (no transfer from 53624)",
        }
    ]
    base_rep = pd.read_csv(rep_path, sep="\t")
    pd.concat([base_rep, pd.DataFrame(append_rows)], ignore_index=True).to_csv(rep_path, sep="\t", index=False)


if __name__ == "__main__":
    main()
