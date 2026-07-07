from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Literal

import numpy as np
import pandas as pd
from lifelines import CoxPHFitter


def _bh_fdr(p: np.ndarray) -> np.ndarray:
    p = np.asarray(p, dtype=float)
    n = np.sum(np.isfinite(p))
    out = np.full_like(p, np.nan)
    if n == 0:
        return out
    idx = np.where(np.isfinite(p))[0]
    pv = p[idx]
    order = np.argsort(pv)
    ranked = pv[order]
    q = ranked * n / (np.arange(len(ranked)) + 1)
    q = np.minimum.accumulate(q[::-1])[::-1]
    out[idx[order]] = np.clip(q, 0, 1)
    return out


@dataclass(frozen=True)
class InteractionConfig:
    endpoint: Literal["OS", "DFS"] = "OS"
    min_n: int = 60
    min_events: int = 15
    min_mutated: int = 8


def screen_mutation_histology_interactions(
    master: pd.DataFrame,
    mut_mat: pd.DataFrame,
    genes: Iterable[str],
    config: InteractionConfig = InteractionConfig(),
    *,
    time_col: str | None = None,
    event_col: str | None = None,
) -> pd.DataFrame:
    if time_col is not None and event_col is not None:
        Tcol, Ecol = time_col, event_col
    elif config.endpoint == "OS":
        Tcol, Ecol = "OS_MONTHS", "OS_EVENT"
    else:
        Tcol, Ecol = "DFS_MONTHS", "DFS_EVENT"

    df0 = master[["SAMPLE_ID", "HISTOLOGY", Tcol, Ecol]].dropna().copy()
    df0 = df0[df0["SAMPLE_ID"].isin(mut_mat.index)]
    if len(df0) < config.min_n or df0[Ecol].sum() < config.min_events:
        return pd.DataFrame()

    df0["HISTOLOGY_BIN"] = (df0["HISTOLOGY"] == "ESCC").astype("int64")

    rows: list[dict] = []
    for g in genes:
        if g not in mut_mat.columns:
            continue
        m = mut_mat.loc[df0["SAMPLE_ID"], g].astype("int64").to_numpy()
        if int(m.sum()) < config.min_mutated:
            continue
        df = df0.copy()
        df["M"] = m
        df["T"] = df[Tcol].astype(float)
        df["E"] = df[Ecol].astype(int)
        # Interaction: hazard ~ M + hist + M:hist
        cph = CoxPHFitter()
        try:
            cph.fit(df[["T", "E", "M", "HISTOLOGY_BIN"]], duration_col="T", event_col="E", formula="M + HISTOLOGY_BIN + M:HISTOLOGY_BIN")
        except Exception:
            continue
        term = "M:HISTOLOGY_BIN"
        if term not in cph.summary.index:
            continue
        rows.append(
            {
                "gene": g,
                "n": int(len(df)),
                "events": int(df["E"].sum()),
                "mutated_n": int(df["M"].sum()),
                "beta_interaction": float(cph.params_.get(term, np.nan)),
                "p_interaction": float(cph.summary.loc[term, "p"]),
                "beta_M_main": float(cph.params_.get("M", np.nan)),
                "p_M_main": float(cph.summary.loc["M", "p"]) if "M" in cph.summary.index else np.nan,
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out["q_interaction"] = _bh_fdr(out["p_interaction"].to_numpy())
    out = out.sort_values("p_interaction").reset_index(drop=True)
    return out

