"""
Shared matplotlib settings for Cancer Medicine manuscript figures: Times New Roman (serif),
300 DPI raster output, strong text contrast.
"""
from __future__ import annotations

from pathlib import Path

import matplotlib as mpl

# Publication raster output (Word/PDF embedding).
PUBLICATION_SAVEFIG_DPI = 300


def apply_publication_style() -> None:
    """Register Times New Roman (with fallbacks) and readable defaults. Call before creating figures."""
    mpl.rcParams.update(
        {
            "font.family": "serif",
            "font.serif": ["Times New Roman", "Times", "Nimbus Roman", "DejaVu Serif"],
            "font.size": 11,
            "axes.titlesize": 12,
            "axes.labelsize": 12,
            "xtick.labelsize": 11,
            "ytick.labelsize": 11,
            "legend.fontsize": 10,
            "figure.titlesize": 13,
            "axes.titleweight": "normal",
            "axes.labelcolor": "black",
            "axes.edgecolor": "#374151",
            "xtick.color": "black",
            "ytick.color": "black",
            "text.color": "black",
            "axes.linewidth": 1.1,
            "grid.color": "#d1d5db",
            "grid.linewidth": 0.8,
            "grid.alpha": 0.95,
            "lines.linewidth": 2.0,
            "lines.antialiased": True,
            "patch.linewidth": 0.6,
            "figure.dpi": 120,
            "savefig.dpi": PUBLICATION_SAVEFIG_DPI,
            "savefig.facecolor": "white",
            "savefig.edgecolor": "none",
            "savefig.bbox": "tight",
            # Math text harmonised with serif body copy where used
            "mathtext.fontset": "stix",
        }
    )


def save_publication_figure(fig: mpl.figure.Figure, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(
        path,
        dpi=PUBLICATION_SAVEFIG_DPI,
        bbox_inches="tight",
        facecolor="white",
        edgecolor="none",
    )
