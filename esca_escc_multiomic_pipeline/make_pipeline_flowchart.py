"""
Generate a clean, publication-ready analysis-pipeline flowchart.
Output: Plots/00_analysis_pipeline_flowchart.png (+ .pdf)
"""
from __future__ import annotations

from esca_escc_multiomic_pipeline.study_paths import study_root

import textwrap
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Circle, FancyArrowPatch, FancyBboxPatch, Polygon, Rectangle, Wedge
import numpy as np

STUDY_ROOT = study_root()
OUT_DIR = STUDY_ROOT / "Plots"

# Layout
FIG_W = 20.0
ROW_GAP = 0.65
SECTION_H = 2.72
BOX_H = 2.60
BOX_Y_OFF = 0.06
MARGIN_Y = 0.15
PAD = 0.08
ICON_FRAC = 0.30
TITLE_BAND = 0.34
ICON_BOX_H_FRAC = 0.98
ICON_SCALE = 1.72

# Vertical stack: E -> D -> C -> B -> A (bottom-up)
Y_E = MARGIN_Y
Y_D = Y_E + SECTION_H + ROW_GAP
Y_C = Y_D + SECTION_H + ROW_GAP
Y_B = Y_C + SECTION_H + ROW_GAP
Y_A = Y_B + SECTION_H + ROW_GAP
FIG_H = Y_A + BOX_Y_OFF + BOX_H + MARGIN_Y

FS_MAIN = 22
FS_SUB = 14
FS_SECTION = 18
FS_TITLE = 18.5
FS_BODY = 16.5
FS_ICON_SM = 11.0
FS_ICON_MD = 12.0
FS_ICON_LG = 13.0

SECTION_STYLES = {
    "A": ("#dbeafe", "#1e40af", "#eff6ff"),
    "B": ("#ede9fe", "#6d28d9", "#f5f3ff"),
    "C": ("#ffe4e6", "#be123c", "#fff1f2"),
    "D": ("#d1fae5", "#047857", "#ecfdf5"),
    "E": ("#ffedd5", "#c2410c", "#fff7ed"),
}


def _sec(section: str):
    return SECTION_STYLES[section]


def _wrap(text: str, width: int) -> str:
    lines: list[str] = []
    for raw in text.split("\n"):
        if not raw.strip():
            lines.append("")
            continue
        if raw.startswith("•"):
            lines.extend(textwrap.wrap(raw, width=width) or [raw])
        else:
            lines.extend(textwrap.wrap(raw, width=width) or [raw])
    return "\n".join(lines)


def _bullet_items(body: str) -> list[str]:
    items: list[str] = []
    for raw in body.split("\n"):
        raw = raw.strip()
        if not raw:
            continue
        if raw.startswith("•"):
            raw = raw[1:].strip()
        items.append(f"• {raw}")
    return items


def _line_height(fontsize: float, spacing: float = 1.12) -> float:
    return fontsize * 0.014 * spacing


def _wrap_item(item: str, wrap_w: int, nowrap: bool = False) -> str:
    if nowrap:
        return item
    return _wrap(item, wrap_w)


def _wrap_chars(width: float, fontsize: float) -> int:
    return max(8, int(width / (fontsize * 0.0078)))


def _count_wrapped_lines(text: str, width: int) -> int:
    return max(1, _wrap(text, width).count("\n") + 1)


def _col_height(items: list[str], wrap_w: int, line_h: float) -> float:
    if not items:
        return 0.0
    lines = sum(_count_wrapped_lines(item, wrap_w) for item in items)
    gaps = max(0, len(items) - 1) * 0.22
    return (lines + gaps) * line_h


def _pick_ncol(n_items: int, text_w: float) -> int:
    if n_items >= 4 and text_w >= 4.0:
        return 2
    return 1


def _clip_text(ax, text_obj, x0: float, y0: float, w: float, h: float):
    clip = Rectangle((x0, y0), w, h, transform=ax.transData, fill=False, visible=False)
    ax.add_patch(clip)
    text_obj.set_clip_path(clip)
    text_obj.set_clip_on(True)


def _apply_clip(artist, clip):
    if clip is not None:
        artist.set_clip_path(clip)
        artist.set_clip_on(True)


def _fit_fontsize(base: float, min_size: float, fits_fn, max_scale: float = 1.20) -> float:
    for scale in (max_scale, 1.16, 1.12, 1.08, 1.04, 1.0):
        size = base * scale
        if fits_fn(size):
            return size
    size = base
    for _ in range(8):
        if fits_fn(size):
            return size
        size = max(min_size, size * 0.92)
    return size


def _split_columns(items: list[str], ncol: int) -> list[list[str]]:
    if ncol <= 1 or len(items) <= 1:
        return [items]
    n = len(items)
    ncols = min(ncol, n)
    base, extra = divmod(n, ncols)
    cols: list[list[str]] = []
    idx = 0
    for c in range(ncols):
        take = base + (1 if c < extra else 0)
        cols.append(items[idx: idx + take])
        idx += take
    return cols


def _anchor(rect, loc: str):
    x, y, w, h = rect
    if loc == "top":
        return (x + w / 2, y + h)
    if loc == "bottom":
        return (x + w / 2, y)
    if loc == "left":
        return (x, y + h / 2)
    if loc == "right":
        return (x + w, y + h / 2)
    return (x + w / 2, y + h / 2)


def _arrow(ax, p1, p2, color="#475569", rad=0.0, lw=1.6):
    style = f"arc3,rad={rad}" if rad else "arc3,rad=0"
    ax.add_patch(
        FancyArrowPatch(
            p1, p2,
            arrowstyle="-|>",
            mutation_scale=14,
            linewidth=lw,
            color=color,
            connectionstyle=style,
            zorder=3,
            shrinkA=4,
            shrinkB=4,
        )
    )


def _draw_icon(ax, kind: str, x0: float, y0: float, x1: float, y1: float,
               clip=None, scale_mul: float = 1.0):
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    w, h = x1 - x0, y1 - y0
    s_w = w * ICON_SCALE
    s_h = h * ICON_SCALE
    s = min(s_w, s_h)
    if h > w * 1.05 and s * 0.62 < h * 0.86:
        s = min(h * ICON_SCALE / 0.66, s_w * 1.06)
    s *= scale_mul

    def line(xs, ys, **kw):
        ln, = ax.plot(xs, ys, clip_on=True, zorder=5, **kw)
        _apply_clip(ln, clip)

    def rect(xy, ww, hh, **kw):
        p = Rectangle(xy, ww, hh, clip_on=True, zorder=5, **kw)
        ax.add_patch(p)
        _apply_clip(p, clip)

    def circ(center, r, **kw):
        p = Circle(center, r, clip_on=True, zorder=5, **kw)
        ax.add_patch(p)
        _apply_clip(p, clip)

    def txt(x, y, s, **kw):
        t = ax.text(x, y, s, **kw)
        _apply_clip(t, clip)
        return t

    if kind == "tcga":
        layers = [
            (cy + 0.20 * s, "#fecaca", "Clinical"),
            (cy + 0.04 * s, "#bbf7d0", "RNA"),
            (cy - 0.12 * s, "#bfdbfe", "Methyl"),
            (cy - 0.28 * s, "#fde68a", "MAF"),
        ]
        for yy, col, lab in layers:
            rect((cx - 0.30 * s, yy - 0.055 * s), 0.60 * s, 0.11 * s, facecolor=col, edgecolor="#334155", lw=0.8)
            txt(cx, yy, lab, ha="center", va="center", fontsize=FS_ICON_MD, fontweight="bold", color="#0f172a", zorder=6)

    elif kind == "harmonize":
        circ((cx - 0.18 * s, cy + 0.10 * s), 0.17 * s, facecolor="#fca5a5", edgecolor="#b91c1c", lw=1.2)
        circ((cx + 0.18 * s, cy + 0.10 * s), 0.17 * s, facecolor="#93c5fd", edgecolor="#1d4ed8", lw=1.2)
        txt(cx - 0.18 * s, cy + 0.10 * s, "EAC", ha="center", va="center", fontsize=FS_ICON_LG, fontweight="bold", color="#7f1d1d", zorder=6)
        txt(cx + 0.18 * s, cy + 0.10 * s, "ESCC", ha="center", va="center", fontsize=FS_ICON_LG, fontweight="bold", color="#1e3a8a", zorder=6)
        poly = Polygon(
            [(cx - 0.24 * s, cy - 0.02 * s), (cx + 0.24 * s, cy - 0.02 * s),
             (cx + 0.08 * s, cy - 0.28 * s), (cx - 0.08 * s, cy - 0.28 * s)],
            closed=True, facecolor="#ddd6fe", edgecolor="#6d28d9", lw=1.1, zorder=5,
        )
        ax.add_patch(poly)
        _apply_clip(poly, clip)
        circ((cx, cy - 0.34 * s), 0.07 * s, facecolor="#7c3aed", edgecolor="#4c1d95", lw=0.9)

    elif kind == "geo":
        rect((cx - 0.32 * s, cy - 0.10 * s), 0.64 * s, 0.40 * s, facecolor="#1e293b", edgecolor="#0f172a", lw=1.2)
        rng = np.random.default_rng(7)
        for r in range(5):
            for c in range(9):
                v = rng.uniform(0, 1)
                col = "#ef4444" if v > 0.75 else "#22c55e" if v > 0.45 else "#eab308" if v > 0.2 else "#475569"
                rect((cx - 0.29 * s + c * 0.064 * s, cy - 0.07 * s + r * 0.072 * s), 0.058 * s, 0.064 * s, facecolor=col, edgecolor="none")
        rect((cx - 0.22 * s, cy - 0.30 * s), 0.44 * s, 0.14 * s, facecolor="#fdba74", edgecolor="#c2410c", lw=1)
        txt(cx, cy - 0.23 * s, "NCBI GEO", ha="center", va="center", fontsize=FS_ICON_MD, color="#7c2d12", fontweight="bold", zorder=6)

    elif kind == "clock":
        circ((cx, cy), 0.34 * s, facecolor="#ffffff", edgecolor="#6d28d9", lw=2)
        wedge = Wedge((cx, cy), 0.34 * s, 90, 270, width=0.08 * s, facecolor="#ede9fe", edgecolor="none", zorder=5)
        ax.add_patch(wedge)
        _apply_clip(wedge, clip)
        for deg in range(0, 360, 30):
            rad = np.deg2rad(deg)
            line([cx + 0.25 * s * np.sin(rad), cx + 0.30 * s * np.sin(rad)],
                 [cy + 0.25 * s * np.cos(rad), cy + 0.30 * s * np.cos(rad)], color="#6d28d9", lw=1.1)
        line([cx, cx], [cy, cy + 0.22 * s], color="#6d28d9", lw=2.2)
        line([cx, cx + 0.14 * s], [cy, cy], color="#8b5cf6", lw=1.8)
        txt(cx, cy - 0.32 * s, "60\u00a0mo.", ha="center", va="top", fontsize=FS_ICON_SM, color="#6d28d9", fontweight="bold", zorder=6)

    elif kind == "matrix":
        pad = 0.008
        bx, by = x0 + pad, y0 + pad
        bw_m, bh_m = w - 2 * pad, h - 2 * pad
        rect((bx, by), bw_m, bh_m, facecolor="#ffffff", edgecolor="#0e7490", lw=1.8)
        half = bw_m / 2
        hdr = 0.24 * bh_m
        hdr_y = by + bh_m - hdr
        rect((bx, hdr_y), half, hdr, facecolor="#0369a1", edgecolor="#0e7490", lw=0.8)
        rect((bx + half, hdr_y), half, hdr, facecolor="#0891b2", edgecolor="#0e7490", lw=0.8)
        line([bx + half, bx + half], [by, hdr_y], color="#0e7490", lw=1.0)
        txt(bx + half * 0.5, hdr_y + hdr / 2, "expr", ha="center", va="center",
            fontsize=FS_ICON_MD, color="white", fontweight="bold", zorder=6)
        txt(bx + half + half * 0.5, hdr_y + hdr / 2, "meth", ha="center", va="center",
            fontsize=FS_ICON_MD, color="white", fontweight="bold", zorder=6)
        expr_vals = [0.85, 0.25, 0.7, 0.15, 0.9, 0.45]
        meth_vals = [0.8, 0.9, 0.35, 0.6, 0.2, 0.75]
        body_h = bh_m - hdr
        row_h = body_h / 6
        for r in range(6):
            yy = by + r * row_h
            ex, me = expr_vals[r], meth_vals[r]
            rect((bx + 0.02, yy), half - 0.04, row_h * 0.88,
                 facecolor="#ef4444" if ex > 0.5 else "#3b82f6", edgecolor="#0e7490", lw=0.5)
            rect((bx + half + 0.02, yy), half - 0.04, row_h * 0.88,
                 facecolor="#16a34a" if me > 0.5 else "#86efac", edgecolor="#0e7490", lw=0.5)

    elif kind == "fast_death":
        xs = np.linspace(cx - 0.36 * s, cx + 0.36 * s, 60)
        dens = np.exp(-((xs - (cx - 0.12 * s)) / (0.18 * s)) ** 2) + 0.35 * np.exp(-((xs - (cx + 0.22 * s)) / (0.07 * s)) ** 2)
        ys = y0 + 0.08 * s + dens / dens.max() * (h - 0.16 * s)
        line(xs, ys, color="#9d174d", lw=2.4)
        rect((cx + 0.10 * s, y0 + 0.08 * s), 0.28 * s, h - 0.16 * s, facecolor="#fecdd3", edgecolor="none", alpha=0.5)
        line([cx + 0.10 * s, cx + 0.10 * s], [y0 + 0.08 * s, y1 - 0.08 * s], color="#be123c", lw=1.6, ls="--")

    elif kind == "mutation":
        t = np.linspace(0, 4 * np.pi, 200)
        rx, ry = 0.40 * s, 0.38 * s
        line(cx + rx * np.sin(t), cy + ry * np.cos(t) * 0.55, color="#166534", lw=3.2)
        for ang in np.linspace(0, 2 * np.pi, 6, endpoint=False):
            px = cx + rx * np.sin(ang)
            py = cy + ry * np.cos(ang) * 0.55
            circ((px, py), 0.055 * s, facecolor="#ef4444", edgecolor="#991b1b", lw=0.8)

    elif kind == "elastic_net":
        xs = np.linspace(cx - 0.32 * s, cx + 0.32 * s, 80)
        coef = 0.40 * s * np.exp(-((xs - (cx - 0.08 * s)) / (0.13 * s)) ** 2) + 0.18 * s * np.exp(-((xs - (cx + 0.16 * s)) / (0.09 * s)) ** 2)
        base = cy - 0.24 * s
        line([cx - 0.32 * s, cx + 0.32 * s], [base, base], color="#64748b", lw=0.9)
        line([cx - 0.32 * s, cx - 0.32 * s], [base, base + 0.42 * s], color="#64748b", lw=0.9)
        line(xs, base + coef, color="#991b1b", lw=2.4)

    elif kind == "cindex":
        t1 = np.linspace(cx - 0.30 * s, cx + 0.30 * s, 50)
        s1 = cy - 0.18 * s + 0.36 * s * np.exp(-0.5 * ((t1 - (cx - 0.08 * s)) / (0.16 * s)) ** 2)
        s2 = cy - 0.18 * s + 0.24 * s * np.exp(-0.5 * ((t1 - (cx + 0.10 * s)) / (0.15 * s)) ** 2)
        line(t1, s1, color="#991b1b", lw=2.2)
        line(t1, s2, color="#64748b", lw=2, ls="--")
        arr = FancyArrowPatch((cx - 0.05 * s, cy + 0.08 * s), (cx + 0.12 * s, cy + 0.20 * s),
                              arrowstyle="<->", mutation_scale=12, color="#991b1b", lw=1.4, zorder=5)
        ax.add_patch(arr)
        _apply_clip(arr, clip)

    elif kind == "geo_cox":
        fs_id = FS_ICON_SM
        fs_sub = FS_ICON_SM - 1.0
        hm_h = h * 0.20
        hm_y = y1 - hm_h - h * 0.03
        rect((x0 + w * 0.07, hm_y), w * 0.86, hm_h, facecolor="#1e293b", edgecolor="#0f172a", lw=1.0)
        rng = np.random.default_rng(7)
        cell_w = w * 0.86 / 9
        cell_h = hm_h / 5
        for r in range(5):
            for c in range(9):
                v = rng.uniform(0, 1)
                col = "#ef4444" if v > 0.75 else "#22c55e" if v > 0.45 else "#eab308" if v > 0.2 else "#475569"
                rect((x0 + w * 0.07 + c * cell_w, hm_y + r * cell_h), cell_w * 0.92, cell_h * 0.88,
                     facecolor=col, edgecolor="none")
        txt((x0 + x1) / 2, hm_y + hm_h / 2, "GEO", ha="center", va="center",
            fontsize=FS_ICON_SM - 1.0, color="#fdba74", fontweight="bold", zorder=6)

        flow_h = hm_y - y0 - h * 0.04
        cy_flow = y0 + flow_h * 0.52
        span = w * 0.88
        bw, bh = span * 0.18, flow_h * 0.62
        cox_r = min(s * 0.050, span * 0.055)
        left_cx = cx - span * 0.36
        right_cx = cx + span * 0.36
        id_y = y1 - hm_h - h * 0.06
        rect((left_cx - bw / 2, cy_flow - bh / 2), bw, bh, facecolor="#fdba74", edgecolor="#c2410c", lw=1.1)
        txt(left_cx, id_y, "GSE53624", ha="center", va="top",
            fontsize=fs_id, fontweight="bold", color="#7c2d12", zorder=6)
        txt(left_cx, cy_flow - bh * 0.20, "train", ha="center", va="center",
            fontsize=fs_sub, color="#9a3412", zorder=6)
        circ((cx, cy_flow), cox_r, facecolor="#fecaca", edgecolor="#991b1b", lw=1.1)
        txt(cx, cy_flow, "Cox", ha="center", va="center", fontsize=fs_id, fontweight="bold", color="#991b1b", zorder=6)
        rect((right_cx - bw / 2, cy_flow - bh / 2), bw, bh, facecolor="#ffedd5", edgecolor="#c2410c", lw=1.1)
        txt(right_cx, id_y, "GSE53622", ha="center", va="top",
            fontsize=fs_id, fontweight="bold", color="#7c2d12", zorder=6)
        txt(right_cx, cy_flow - bh * 0.20, "hold-out", ha="center", va="center",
            fontsize=fs_sub, color="#9a3412", zorder=6)
        for p1, p2 in [((left_cx + bw / 2, cy_flow), (cx - cox_r, cy_flow)),
                       ((cx + cox_r, cy_flow), (right_cx - bw / 2, cy_flow))]:
            arr = FancyArrowPatch(p1, p2, arrowstyle="-|>", mutation_scale=8, color="#c2410c", lw=1.1, zorder=5)
            ax.add_patch(arr)
            _apply_clip(arr, clip)

    elif kind == "volcano":
        rng = np.random.default_rng(4)
        sc = ax.scatter(rng.normal(cx, 0.11 * s, 90), rng.normal(cy - 0.04 * s, 0.12 * s, 90),
                        s=10, c="#94a3b8", alpha=0.55, linewidths=0, zorder=5)
        _apply_clip(sc, clip)
        wing_x = np.linspace(cx - 0.30 * s, cx + 0.30 * s, 60)
        line(wing_x, cy - 0.04 * s + 0.26 * s * np.abs(wing_x - cx) / (0.30 * s), color="#cbd5e1", lw=1.2)
        sc2 = ax.scatter([cx - 0.16 * s, cx + 0.18 * s], [cy + 0.22 * s, cy + 0.24 * s],
                         s=60, c=["#2563eb", "#dc2626"], edgecolors="#0f172a", linewidths=0.6, zorder=6)
        _apply_clip(sc2, clip)

    elif kind == "jaccard":
        margin = w * 0.08
        span = w - 2 * margin
        positions = [cx - span / 2, cx, cx + span / 2]
        cy_circ = y0 + h * 0.58
        pct_y = y0 + h * 0.12
        for ox, pct in zip(positions, ["10%", "20%", "30%"]):
            r = min(0.065 * s, h * 0.22)
            offset = r * 0.82
            circ((ox - offset, cy_circ), r, facecolor="#0d9488", edgecolor="#115e59", lw=1, alpha=0.85)
            circ((ox + offset, cy_circ), r, facecolor="#5eead4", edgecolor="#115e59", lw=1, alpha=0.85)
            circ((ox, cy_circ), r * 0.36, facecolor="#134e4a", edgecolor="none", alpha=0.55)
            txt(ox, pct_y, pct, ha="center", va="center",
                fontsize=FS_ICON_SM - 1.0, color="#115e59", fontweight="bold", zorder=6)

    elif kind == "cox_univ":
        probes = [cx - 0.24 * s, cx - 0.08 * s, cx + 0.08 * s, cx + 0.24 * s]
        hrs = [0.6, 1.0, 1.7, 2.4]
        ref = cy - 0.06 * s
        line([cx - 0.28 * s, cx + 0.28 * s], [ref, ref], color="#64748b", lw=0.9)
        for p, hr in zip(probes, hrs):
            col = "#2563eb" if hr < 1 else "#dc2626"
            yp = ref + (hr - 1) * 0.11 * s
            line([p, p], [ref, yp], color=col, lw=2)
            circ((p, yp), 0.032 * s, facecolor=col, edgecolor="#0f172a", lw=0.5)

    elif kind == "km":
        t = np.array([0, 15, 15, 30, 30, 45, 45, 60])
        hi = np.array([1.0, 1.0, 0.78, 0.78, 0.55, 0.55, 0.38, 0.38])
        lo = np.array([1.0, 1.0, 0.88, 0.88, 0.72, 0.72, 0.58, 0.58])
        tx = cx - 0.30 * s + t / 60 * 0.60 * s
        base = cy - 0.22 * s
        line(tx, base + hi * 0.44 * s, color="#dc2626", lw=2.2, drawstyle="steps-post")
        line(tx, base + lo * 0.44 * s, color="#2563eb", lw=2.2, drawstyle="steps-post")
        line([cx - 0.30 * s, cx + 0.30 * s], [base, base], color="#64748b", lw=0.8)
        line([cx - 0.30 * s, cx - 0.30 * s], [base, base + 0.46 * s], color="#64748b", lw=0.8)

    elif kind == "interaction":
        xs = np.linspace(cx - 0.28 * s, cx + 0.28 * s, 40)
        line(xs, cy - 0.08 * s + 0.018 * s * (xs - cx) / s, color="#2563eb", lw=2.2)
        line(xs, cy + 0.08 * s - 0.035 * s * (xs - cx) / s, color="#dc2626", lw=2.2)
        txt(cx + 0.04 * s, cy + 0.20 * s, "EAC", fontsize=FS_ICON_MD, color="#2563eb", fontweight="bold", zorder=6)
        txt(cx + 0.04 * s, cy - 0.24 * s, "ESCC", fontsize=FS_ICON_MD, color="#dc2626", fontweight="bold", zorder=6)

    elif kind == "hallmark":
        modules = ["EMT", "Prol.", "Imm.", "IFNγ"]
        mut_h = [0.68, 0.50, 0.34, 0.58]
        wt_h = [0.42, 0.44, 0.44, 0.36]
        margin = w * 0.04
        span = w - 2 * margin
        group_step = span / 3.6
        bar_w = min(0.050 * s, group_step * 0.18)
        pair_gap = 0.008 * s
        base = y0 + 0.22 * s
        max_h = h * 0.48
        label_y = y0 + 0.02 * s
        x_start = cx - span / 2 + bar_w
        for i, (m, mh, wh) in enumerate(zip(modules, mut_h, wt_h)):
            xb = x_start + i * group_step
            rect((xb, base), bar_w, mh * max_h, facecolor="#dc2626", edgecolor="#991b1b", lw=0.7)
            rect((xb + bar_w + pair_gap, base), bar_w, wh * max_h, facecolor="#2563eb", edgecolor="#1d4ed8", lw=0.7)
            label_x = xb + bar_w + pair_gap / 2
            txt(label_x, label_y, m, ha="center", va="bottom",
                fontsize=FS_ICON_SM - 1.0, color="#78350f", fontweight="bold", zorder=6)

    elif kind == "mediation":
        pad_x, pad_y = w * 0.015, h * 0.02
        ux0, uy0, ux1, uy1 = x0 + pad_x, y0 + pad_y, x1 - pad_x, y1 - pad_y
        uw, uh = ux1 - ux0, uy1 - uy0
        s_m = min(uw / 0.80, uh / 0.74)
        cx_m = (ux0 + ux1) / 2
        cy_m = uy0 + uh * 0.44
        bw, bh = 0.21 * s_m, 0.12 * s_m
        nodes = [
            (cx_m - 0.31 * s_m, cy_m, "#bfdbfe", "Meth (β)"),
            (cx_m, cy_m + 0.25 * s_m, "#bbf7d0", "RNA"),
            (cx_m + 0.31 * s_m, cy_m, "#fecaca", "OS"),
        ]
        for bx, by, col, lab in nodes:
            node = FancyBboxPatch((bx - bw / 2, by - bh / 2), bw, bh, boxstyle="round,pad=0.02",
                                  facecolor=col, edgecolor="#581c87", lw=1.2, zorder=5)
            ax.add_patch(node)
            _apply_clip(node, clip)
            txt(bx, by, lab, ha="center", va="center", fontsize=FS_ICON_SM, fontweight="bold",
                color="#1e293b", zorder=6)
        for p1, p2, dashed in [
            ((cx_m - 0.21 * s_m, cy_m + 0.10 * s_m), (cx_m - 0.05 * s_m, cy_m + 0.18 * s_m), False),
            ((cx_m + 0.05 * s_m, cy_m + 0.18 * s_m), (cx_m + 0.21 * s_m, cy_m + 0.10 * s_m), False),
            ((cx_m - 0.21 * s_m, cy_m - 0.14 * s_m), (cx_m + 0.21 * s_m, cy_m - 0.14 * s_m), True),
        ]:
            if dashed:
                arr = FancyArrowPatch(
                    p1, p2, arrowstyle="-|>", mutation_scale=10, color="#a21caf", lw=1.2,
                    linestyle="dashed", connectionstyle="arc3,rad=-0.18", zorder=5,
                )
            else:
                arr = FancyArrowPatch(
                    p1, p2, arrowstyle="-|>", mutation_scale=10, color="#7e22ce", lw=1.4, zorder=5,
                )
            ax.add_patch(arr)
            _apply_clip(arr, clip)
        txt(cx_m - 0.14 * s_m, cy_m + 0.21 * s_m, "a", fontsize=FS_ICON_SM,
            color="#7e22ce", fontweight="bold", zorder=6)
        txt(cx_m + 0.14 * s_m, cy_m + 0.21 * s_m, "b", fontsize=FS_ICON_SM,
            color="#7e22ce", fontweight="bold", zorder=6)
        txt(cx_m, cy_m - 0.28 * s_m, "a×b", ha="center", fontsize=FS_ICON_SM,
            color="#86198f", fontweight="bold", zorder=6)


def _section_band(ax, y: float, h: float, section: str, label: str):
    fill, edge, bg = _sec(section)
    band_h = 0.55
    band_top = y + h - band_h
    ax.add_patch(Rectangle((0.35, y), FIG_W - 0.7, h, facecolor=bg, edgecolor="none", zorder=0))
    ax.add_patch(Rectangle((0.35, band_top), FIG_W - 0.7, band_h, facecolor=edge, edgecolor="none", zorder=1, alpha=0.12))
    ax.text(0.42, y + h - 0.06, section, fontsize=FS_SECTION + 3, fontweight="bold",
            color=edge, va="top", ha="left", zorder=2)
    title = label.lstrip(section).strip()
    title_wrap = _wrap_chars(4.8, FS_SECTION)
    title_wrapped = _wrap(title, title_wrap)
    title_obj = ax.text(
        0.82, band_top + band_h / 2, title_wrapped,
        fontsize=FS_SECTION, fontweight="bold", color=edge,
        va="center", ha="left", zorder=2, linespacing=1.05,
    )
    title_clip = Rectangle((0.35, band_top), 4.65, band_h, transform=ax.transData, fill=False, visible=False)
    ax.add_patch(title_clip)
    _apply_clip(title_obj, title_clip)


def _box(ax, x: float, y: float, w: float, h: float, title: str, body: str, section: str,
         icon: str | None = None, icon_frac: float | None = None,
         icon_pad_bottom: float = 0.005, icon_extend_left: float = 0.0,
         title_oneline: bool = False, icon_scale_mul: float = 1.0,
         nowrap_body: bool = False, compact: bool = False, icon_inset: float = 0.004,
         icon_position: str = "right", icon_bottom_h_frac: float = 0.38):
    fill, edge, _ = _sec(section)
    patch = FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.10",
        linewidth=2.0, edgecolor=edge, facecolor=fill, zorder=4,
    )
    ax.add_patch(patch)
    ax.add_patch(Rectangle((x + 0.05, y + 0.03), 0.10, h - 0.06, facecolor=edge, edgecolor="none", zorder=5, alpha=0.85))

    clip_x = x + 0.10
    clip_y = y + 0.03
    clip_w = w - 0.20
    clip_h = h - 0.06

    box_clip = Rectangle((x + 0.02, y + 0.015), w - 0.04, h - 0.03, transform=ax.transData, fill=False, visible=False)
    ax.add_patch(box_clip)

    inner_x = x + PAD + 0.14
    inner_right = x + w - PAD
    inner_w = inner_right - inner_x
    icon_gap = 0.05
    frac = icon_frac if icon_frac is not None else ICON_FRAC
    icon_w = inner_w * frac if icon else 0.0
    if icon and icon_position == "bottom":
        text_w = inner_w - 0.06
    else:
        text_right = inner_right - icon_w - icon_extend_left - (icon_gap if icon else 0.0)
        text_w = max(0.5, text_right - inner_x - 0.06)

    title_avail_w = inner_w - 0.06 if (title_oneline or icon_position == "bottom") else text_w

    # --- Title (top band, auto-fit) ---
    def title_fits(fs: float) -> bool:
        if title_oneline:
            est_w = len(title) * fs * 0.0078
            return est_w <= title_avail_w and _line_height(fs, 1.05) <= TITLE_BAND - 0.03
        tw = _wrap_chars(title_avail_w, fs)
        lines = _count_wrapped_lines(title, tw)
        return lines <= 2 and lines * _line_height(fs, 1.05) <= TITLE_BAND - 0.03

    title_fs = _fit_fontsize(FS_TITLE, 13.5, title_fits)
    if title_oneline:
        title_wrapped = title
    else:
        title_wrap = _wrap_chars(title_avail_w, title_fs)
        title_wrapped = _wrap(title, title_wrap)
    title_y = y + h - (0.02 if compact else 0.03)
    title_lines = max(1, title_wrapped.count("\n") + 1)
    title_block_h = title_lines * _line_height(title_fs, 1.03) + (0.01 if compact else 0.02)

    title_obj = ax.text(
        inner_x, title_y, title_wrapped,
        ha="left", va="top", fontsize=title_fs, fontweight="bold",
        color="#0f172a", zorder=6, linespacing=1.03,
    )
    _clip_text(ax, title_obj, inner_x - 0.02, y + h - title_block_h - 0.01,
               title_avail_w + 0.04, title_block_h + 0.01)
    _apply_clip(title_obj, box_clip)

    # --- Icon + text zones ---
    body_top = y + h - title_block_h - (0.002 if compact else 0.005)
    body_bottom = y + (0.001 if compact else 0.002)
    body_h = body_top - body_bottom
    text_area_top = body_top
    text_area_bottom = body_bottom

    if icon:
        if icon_position == "bottom":
            icon_box_h = body_h * icon_bottom_h_frac
            iy0 = body_bottom + 0.003
            iy1 = iy0 + icon_box_h
            icon_w_draw = inner_w * frac
            ix0 = inner_x + (inner_w - icon_w_draw) * 0.5
            text_area_bottom = iy1 + 0.010
        else:
            ix0 = inner_right - icon_w - icon_extend_left
            icon_w_draw = icon_w + icon_extend_left
            icon_box_frac = 0.99 if compact else ICON_BOX_H_FRAC
            icon_box_h = min(body_h * icon_box_frac, body_h - (0.003 if compact else 0.006))
            iy0 = body_bottom + (0.002 if compact else 0.004)
            iy1 = iy0 + icon_box_h
            if iy1 > body_top:
                iy1 = body_top
                iy0 = max(body_bottom + 0.004, iy1 - icon_box_h)

        icon_h = max(0.2, iy1 - iy0)
        ax.add_patch(FancyBboxPatch(
            (ix0, iy0), icon_w_draw, icon_h,
            boxstyle="round,pad=0.01,rounding_size=0.06",
            facecolor="#ffffff", edgecolor=edge, linewidth=1.2, alpha=0.92, zorder=5,
        ))
        icon_clip = Rectangle((ix0, iy0), icon_w_draw, icon_h, transform=ax.transData, fill=False, visible=False)
        ax.add_patch(icon_clip)
        icon_pad_bot = max(0.0, icon_pad_bottom * 0.5)
        inset = icon_inset
        _draw_icon(
            ax, icon,
            ix0 + inset, iy0 + icon_pad_bot,
            ix0 + icon_w_draw - inset, iy1 - inset,
            clip=icon_clip,
            scale_mul=icon_scale_mul,
        )

    text_body_h = text_area_top - text_area_bottom
    items = _bullet_items(body)
    ncol = _pick_ncol(len(items), text_w)
    col_gap = 0.12
    col_w = (text_w - col_gap * (ncol - 1)) / ncol

    def body_fits(fs: float) -> bool:
        lh = _line_height(fs)
        wrap = 999 if nowrap_body else _wrap_chars(col_w, fs)
        cols_fit = _split_columns(items, ncol)
        if nowrap_body:
            char_w = fs * 0.0078
            for col in cols_fit:
                for item in col:
                    if len(item) * char_w > col_w + 0.08:
                        return False
            need = sum(len(col) for col in cols_fit) * lh * 1.10
            need += max(0, max((len(c) for c in cols_fit), default=0) - 1) * 0.22 * lh
        else:
            need = max(_col_height(col, wrap, lh) for col in cols_fit)
        return need <= text_body_h - 0.006

    body_fs = _fit_fontsize(FS_BODY, 14.0, body_fits)
    line_h = _line_height(body_fs)
    wrap_w = 999 if nowrap_body else _wrap_chars(col_w, body_fs)
    cols = _split_columns(items, ncol)

    col_heights: list[float] = []
    col_texts: list[str] = []
    for ci, col_items in enumerate(cols):
        col_x = inner_x + ci * (col_w + col_gap)
        col_lines: list[str] = []
        for i, item in enumerate(col_items):
            if i > 0:
                col_lines.append("")
            col_lines.append(_wrap_item(item, wrap_w, nowrap=nowrap_body))
        col_text = "\n".join(col_lines)
        col_texts.append(col_text)
        lines = max(1, col_text.count("\n") + 1)
        gaps = max(0, len(col_items) - 1) * 0.22
        col_heights.append((lines + gaps) * line_h * 1.10)

    text_anchor_y = text_area_top - 0.006

    for ci, col_items in enumerate(cols):
        col_x = inner_x + ci * (col_w + col_gap)
        body_obj = ax.text(
            col_x, text_anchor_y, col_texts[ci],
            ha="left", va="top", fontsize=body_fs, color="#334155",
            zorder=6, linespacing=1.10,
        )
        _clip_text(ax, body_obj, col_x - 0.02, text_area_bottom, col_w + 0.04, text_body_h)
        _apply_clip(body_obj, box_clip)

    return (x, y, w, h)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    plt.rcParams.update({
        "font.family": "serif",
        "font.serif": ["Times New Roman", "Times", "DejaVu Serif"],
        "figure.facecolor": "white",
        "savefig.facecolor": "white",
    })

    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H), facecolor="white")
    ax.set_xlim(0, FIG_W)
    ax.set_ylim(0, FIG_H)
    ax.axis("off")
    ax.set_position([0, 0, 1, 1])

    # --- A: Data sources ---
    yA = Y_A
    hA = SECTION_H
    _section_band(ax, yA, hA, "A", "A   Data sources")
    bw = 5.85
    gap = 0.55
    x0 = 0.75
    g_tcga = _box(ax, x0, yA + BOX_Y_OFF, bw, BOX_H,
                  "TCGA ESCA (Firehose Legacy)",
                  "• cBioPortal download\n• Clinical OS / DFS\n• HM450 methylation (β)\n• RNA z-scores + MAF",
                  "A", icon="tcga")
    g_harm = _box(ax, x0 + bw + gap, yA + BOX_Y_OFF, bw, BOX_H,
                  "Sample harmonisation",
                  "• Primary tumors only\n• EAC (n=89) vs ESCC (n=95)\n• Omics overlap filter\n• OncoTree histology",
                  "A", icon="harmonize")
    g_geo = _box(ax, x0 + 2 * (bw + gap), yA + BOX_Y_OFF, bw + 0.4, BOX_H,
                 "External GEO cohorts",
                 "• GSE161533 — ESCC expression\n• GSE53624 — lncRNA (train)\n• GSE53622 — lncRNA (hold-out)",
                 "A", icon="geo")
    _, ea, _ = _sec("A")
    _arrow(ax, _anchor(g_tcga, "right"), _anchor(g_harm, "left"), color=ea)
    _arrow(ax, _anchor(g_harm, "right"), _anchor(g_geo, "left"), color=ea)

    # --- B: Pre-processing ---
    yB = Y_B
    hB = SECTION_H
    _section_band(ax, yB, hB, "B", "B   Pre-processing & endpoint")
    bwB = 4.35
    g_os60 = _box(ax, x0, yB + BOX_Y_OFF, bwB, BOX_H, "OS60 truncated endpoint",
                  "• Truncate follow-up at 60\u00a0mo.\n• Late events censored", "B", icon="clock",
                  nowrap_body=True)
    g_feat = _box(ax, x0 + bwB + 0.35, yB + BOX_Y_OFF, bwB, BOX_H, "Feature matrix",
                  "• Paired expr\u00a0+\u00a0meth\n• ≤60 genes per histology\n• Median imputation", "B", icon="matrix",
                  nowrap_body=True)
    g_fast = _box(ax, x0 + 2 * (bwB + 0.35), yB + BOX_Y_OFF, bwB + 0.15, BOX_H, "Short-survival phenotype",
                  "• Shortest-survival tail\n• Among OS60 decedents\n• 10–30% (min n=10)", "B", icon="fast_death")
    g_mut = _box(ax, x0 + 3 * (bwB + 0.35) + 0.15, yB + BOX_Y_OFF, bwB + 0.25, BOX_H, "Mutation matrix",
                 "• Top 80 mutated genes\n• Binary calls (≥8 mutants)", "B", icon="mutation")
    _, eb, _ = _sec("B")
    _arrow(ax, _anchor(g_os60, "right"), _anchor(g_feat, "left"), color=eb)
    _arrow(ax, _anchor(g_feat, "right"), _anchor(g_fast, "left"), color=eb)
    _arrow(ax, _anchor(g_fast, "right"), _anchor(g_mut, "left"), color=eb)
    _arrow(ax, _anchor(g_harm, "bottom"), _anchor(g_feat, "top"), color=ea, rad=-0.08)

    # --- C: ML / survival ---
    yC = Y_C
    hC = SECTION_H
    _section_band(ax, yC, hC, "C", "C   Machine-learning / survival models")
    g_cox = _box(ax, x0, yC + BOX_Y_OFF, 6.1, BOX_H,
                 "Elastic-net Cox PH",
                 "• CoxPHFitter: elastic net\n• EAC / ESCC / pooled\n• Fixed-coefficient transfer",
                 "C", icon="elastic_net")
    g_cidx = _box(ax, x0 + 6.45, yC + BOX_Y_OFF, 5.8, BOX_H,
                  "Model performance",
                  "• Harrell C-index\n• Within-/cross-lineage\n• GDC 2025: 5-fold CV", "C", icon="cindex")
    g_geocox = _box(ax, x0 + 12.55, yC + BOX_Y_OFF, 6.0, BOX_H,
                    "GEO penalized Cox (lncRNA)",
                    "• Elastic-net: GSE53624\n• Top-60 lncRNAs\n• Fixed scoring: GSE53622", "C", icon="geo_cox",
                    icon_frac=0.50, icon_extend_left=0.30, title_oneline=True, icon_scale_mul=0.82)
    _, ec, _ = _sec("C")
    _arrow(ax, _anchor(g_cox, "right"), _anchor(g_cidx, "left"), color=ec)
    _arrow(ax, _anchor(g_feat, "bottom"), _anchor(g_cox, "top"), color=eb, rad=-0.06)
    _arrow(ax, _anchor(g_geo, "bottom"), _anchor(g_geocox, "top"), color=ea, rad=0.10)

    # --- D: Biostatistics ---
    yD = Y_D
    hD = SECTION_H
    _section_band(ax, yD, hD, "D", "D   Biostatistical tests & corrections")
    g_volc = _box(ax, x0, yD + BOX_Y_OFF, bwB, BOX_H,
                  "Differential methylation",
                  "• Short-survival vs others\n• Welch t-test (HM450 β)\n• BH-FDR (q_diff)", "D", icon="volcano")
    g_jacc = _box(ax, x0 + bwB + 0.35, yD + BOX_Y_OFF, bwB, BOX_H,
                  "Sensitivity",
                  "• Patient-set Jaccard\n• Probe-set Jaccard\n• 10 / 20 / 30% tails", "D", icon="jaccard",
                  icon_position="bottom", icon_frac=0.90, icon_bottom_h_frac=0.34,
                  icon_scale_mul=0.78, nowrap_body=True, compact=True)
    g_coxu = _box(ax, x0 + 2 * (bwB + 0.35), yD + BOX_Y_OFF, bwB + 0.15, BOX_H,
                  "Methylation vs OS (EAC)",
                  "• Univariate Cox PH (beta)\n• BH-FDR on p-values", "D", icon="cox_univ")
    g_km = _box(ax, x0 + 3 * (bwB + 0.35) + 0.15, yD + BOX_Y_OFF, bwB + 0.25, BOX_H,
                "Kaplan–Meier OS60",
                "• Median β split\n• Log-rank (high vs low)", "D", icon="km")
    _, ed, _ = _sec("D")
    _arrow(ax, _anchor(g_volc, "right"), _anchor(g_jacc, "left"), color=ed)
    _arrow(ax, _anchor(g_jacc, "right"), _anchor(g_coxu, "left"), color=ed)
    _arrow(ax, _anchor(g_coxu, "right"), _anchor(g_km, "left"), color=ed)
    _arrow(ax, _anchor(g_fast, "bottom"), _anchor(g_volc, "top"), color=eb, rad=-0.06)
    _arrow(ax, _anchor(g_cox, "bottom"), _anchor(g_coxu, "top"), color=ec, rad=0.08)

    # --- E: Exploratory ---
    yE = Y_E
    hE = SECTION_H
    _section_band(ax, yE, hE, "E", "E   Exploratory somatic & mediation layers")
    g_int = _box(ax, x0, yE + BOX_Y_OFF, 6.2, BOX_H,
                 "Mutation × histology screen",
                 "• Cox PH: M + histology + M×hist\n• Nominal Wald p (OS60)\n• Top 80 genes (≥8 mutants)", "E", icon="interaction")
    g_mech = _box(ax, x0 + 6.55, yE + BOX_Y_OFF, 5.9, BOX_H,
                  "Mechanistic follow-up",
                  "• MSigDB Hallmark scores\n• EMT, proliferation, immune, IFN-γ\n• Welch t-test, mut\u00a0vs\u00a0WT", "E", icon="hallmark",
                  icon_pad_bottom=0.01, nowrap_body=True)
    g_med = _box(ax, x0 + 12.75, yE + BOX_Y_OFF, 5.8, BOX_H,
                 "Mediation-style scan",
                 "• Methylation -> expression -> OS\n• Indirect effect (a\u00a0×\u00a0b)\n• BH-FDR across genes", "E", icon="mediation",
                 icon_frac=0.40, icon_inset=0.0005, icon_scale_mul=1.08, nowrap_body=True)
    _, ee, _ = _sec("E")
    _arrow(ax, _anchor(g_int, "right"), _anchor(g_mech, "left"), color=ee)
    _arrow(ax, _anchor(g_mech, "right"), _anchor(g_med, "left"), color=ee)
    _arrow(ax, _anchor(g_mut, "bottom"), _anchor(g_int, "top"), color=eb, rad=-0.06)
    _arrow(ax, _anchor(g_volc, "bottom"), _anchor(g_med, "top"), color=ed, rad=0.10)

    png = OUT_DIR / "00_analysis_pipeline_flowchart.png"
    pdf = OUT_DIR / "00_analysis_pipeline_flowchart.pdf"
    fig.savefig(png, dpi=300, bbox_inches="tight", facecolor="white", pad_inches=0.0)
    fig.savefig(pdf, bbox_inches="tight", facecolor="white", pad_inches=0.0)
    plt.close(fig)
    print(f"Wrote {png}")
    print(f"Wrote {pdf}")


if __name__ == "__main__":
    main()
