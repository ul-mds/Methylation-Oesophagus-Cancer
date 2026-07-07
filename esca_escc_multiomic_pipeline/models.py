from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Literal

import numpy as np
import pandas as pd
from lifelines import CoxPHFitter


@dataclass(frozen=True)
class ModelConfig:
    endpoint: Literal["OS", "DFS"] = "OS"
    penalizer: float = 0.05
    l1_ratio: float = 0.5
    min_events: int = 15


def _endpoint_cols(endpoint: str) -> tuple[str, str]:
    if endpoint == "OS":
        return "OS_MONTHS", "OS_EVENT"
    if endpoint == "DFS":
        return "DFS_MONTHS", "DFS_EVENT"
    raise ValueError("endpoint must be OS or DFS")


def build_feature_table(
    master: pd.DataFrame,
    expr: pd.DataFrame,
    meth: pd.DataFrame,
    gene_to_probe: dict[str, str],
    genes: Iterable[str],
) -> pd.DataFrame:
    genes = [g for g in genes if g in expr.index and g in gene_to_probe and gene_to_probe[g] in meth.index]
    samples = master["SAMPLE_ID"].astype(str).tolist()
    x_expr = expr.loc[genes, samples].T
    x_expr.columns = [f"expr::{g}" for g in genes]
    x_meth = meth.loc[[gene_to_probe[g] for g in genes], samples].T
    x_meth.columns = [f"meth::{g}" for g in genes]
    X = pd.concat([x_expr, x_meth], axis=1)
    X.index.name = "SAMPLE_ID"
    return X


def fit_cox(
    df: pd.DataFrame,
    duration_col: str,
    event_col: str,
    config: ModelConfig,
    strata: list[str] | None = None,
) -> CoxPHFitter:
    cph = CoxPHFitter(penalizer=config.penalizer, l1_ratio=config.l1_ratio)
    cph.fit(df, duration_col=duration_col, event_col=event_col, strata=strata)
    return cph


def score_risk(cph: CoxPHFitter, X: pd.DataFrame) -> pd.Series:
    # higher = higher risk
    s = cph.predict_partial_hazard(X).iloc[:, 0]
    s.name = "risk_score"
    return s


def concordance(cph: CoxPHFitter) -> float:
    return float(cph.concordance_index_)

