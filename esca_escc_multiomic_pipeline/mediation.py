from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np
import pandas as pd
from lifelines import CoxPHFitter
from scipy.stats import spearmanr


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
class MediationConfig:
    endpoint: Literal["OS", "DFS"] = "OS"
    min_n: int = 35
    # we prefilter to speed up: require some methyl-expression association
    prefilter_abs_spearman: float = 0.25
    # require some survival signal in the Cox step
    max_p_expr_in_cox: float = 0.10


def _cox_expr_given_meth(
    durations: pd.Series,
    events: pd.Series,
    expr: pd.Series,
    meth: pd.Series,
) -> tuple[float, float]:
    # Skip degenerate predictors (helps convergence and speed)
    if float(pd.Series(meth).var(skipna=True)) < 1e-8 or float(pd.Series(expr).var(skipna=True)) < 1e-8:
        return np.nan, np.nan
    df = pd.DataFrame(
        {
            "T": durations.astype(float),
            "E": events.astype(int),
            "expr": expr.astype(float),
            "meth": meth.astype(float),
        }
    ).dropna()
    if len(df) < 10 or df["E"].sum() < 5:
        return np.nan, np.nan
    cph = CoxPHFitter(penalizer=0.1)
    cph.fit(df, duration_col="T", event_col="E", formula="expr + meth")
    b = float(cph.params_.get("expr", np.nan))
    p = float(cph.summary.loc["expr", "p"])
    return b, p


def _linear_expr_on_meth(expr: pd.Series, meth: pd.Series) -> tuple[float, float]:
    df = pd.DataFrame({"expr": expr.astype(float), "meth": meth.astype(float)}).dropna()
    if len(df) < 10:
        return np.nan, np.nan
    x = df["meth"].to_numpy()
    y = df["expr"].to_numpy()
    x = x - x.mean()
    denom = float(np.sum(x**2))
    if denom <= 0:
        return np.nan, np.nan
    a = float(np.sum(x * (y - y.mean())) / denom)  # slope
    # simple t-test via correlation approximation
    r = np.corrcoef(df["meth"], df["expr"])[0, 1]
    if not np.isfinite(r):
        return a, np.nan
    n = len(df)
    t = r * np.sqrt((n - 2) / max(1e-12, 1 - r**2))
    # two-sided p from normal approx for speed
    from scipy.stats import t as tdist

    p = float(2 * tdist.sf(np.abs(t), df=n - 2))
    return a, p


def mediation_scan_by_gene(
    master: pd.DataFrame,
    expr: pd.DataFrame,
    meth: pd.DataFrame,
    gene_to_probe: dict[str, str],
    histology_label: str,
    config: MediationConfig = MediationConfig(),
) -> pd.DataFrame:
    """
    For each gene (with a chosen representative probe), estimate:
    - a: methylation -> expression (linear slope)
    - b: expression -> hazard in Cox, controlling for methylation
    - indirect proxy: a*b

    This is a pragmatic survival-mediation approximation (product-of-coefficients with Cox for outcome).
    """
    if config.endpoint == "OS":
        Tcol, Ecol = "OS_MONTHS", "OS_EVENT"
    else:
        Tcol, Ecol = "DFS_MONTHS", "DFS_EVENT"

    sub = master[master["HISTOLOGY"] == histology_label].copy()
    sub = sub[[Tcol, Ecol, "SAMPLE_ID"]].dropna()
    if len(sub) < config.min_n:
        return pd.DataFrame()

    # Crucial: index survival vectors by SAMPLE_ID to align with expr/meth series
    sub["SAMPLE_ID"] = sub["SAMPLE_ID"].astype(str)
    sub = sub.set_index("SAMPLE_ID", drop=True)
    samples = sub.index.tolist()
    durations = sub[Tcol].astype(float)
    events = sub[Ecol].astype(int)

    rows: list[dict] = []
    for gene, probe in gene_to_probe.items():
        if gene not in expr.index or probe not in meth.index:
            continue
        y_expr = expr.loc[gene, samples]
        x_meth = meth.loc[probe, samples]
        # Ensure sample-indexed vectors (avoid mixed int/str indexes in downstream DataFrames)
        y_expr.index = y_expr.index.astype(str)
        x_meth.index = x_meth.index.astype(str)
        # prefilter on methyl-expr association
        ok = pd.concat([y_expr, x_meth], axis=1).dropna()
        if len(ok) < config.min_n:
            continue
        rho, rho_p = spearmanr(ok.iloc[:, 0].to_numpy(), ok.iloc[:, 1].to_numpy())
        if not np.isfinite(rho) or abs(rho) < config.prefilter_abs_spearman:
            continue

        a, a_p = _linear_expr_on_meth(y_expr, x_meth)
        # Align durations/events to the same sample index as expression/methylation
        idx = ok.index.astype(str)
        b, b_p = _cox_expr_given_meth(durations.loc[idx], events.loc[idx], y_expr.loc[idx], x_meth.loc[idx])
        if not np.isfinite(b_p) or b_p > config.max_p_expr_in_cox:
            continue

        rows.append(
            {
                "gene": gene,
                "probe": probe,
                "n": int(len(ok)),
                "spearman_rho": float(rho),
                "spearman_p": float(rho_p),
                "a_meth_to_expr": float(a),
                "a_p": float(a_p),
                "b_expr_in_cox": float(b),
                "b_p": float(b_p),
                "indirect_ab": float(a * b) if np.isfinite(a) and np.isfinite(b) else np.nan,
                "histology": histology_label,
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out["a_q"] = _bh_fdr(out["a_p"].to_numpy())
    out["b_q"] = _bh_fdr(out["b_p"].to_numpy())
    out = out.sort_values(["b_p", "a_p"], ascending=[True, True]).reset_index(drop=True)
    return out

