"""60-month landmark truncation for OS (TCGA-style)."""
from __future__ import annotations

import numpy as np
import pandas as pd


def truncate_os_at_months(
    os_months: pd.Series,
    os_event: pd.Series,
    months: float = 60.0,
) -> tuple[pd.Series, pd.Series]:
    """
    Landmark analysis: follow-up capped at `months`.
    Event = 1 only if death occurred at or before `months`.
    """
    t = pd.to_numeric(os_months, errors="coerce")
    e = pd.to_numeric(os_event, errors="coerce").fillna(0).astype(int)
    t_trunc = np.minimum(t.to_numpy(dtype=float), float(months))
    e_trunc = np.where((e.to_numpy(dtype=int) == 1) & (t.to_numpy(dtype=float) <= float(months)), 1, 0)
    return pd.Series(t_trunc, index=os_months.index), pd.Series(e_trunc, index=os_months.index)
