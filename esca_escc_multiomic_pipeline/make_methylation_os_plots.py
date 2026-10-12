from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

try:
    from .plot_style import apply_publication_style, bolden_axis_text, save_publication_figure
    from .survival_60mo import truncate_os_at_months
except ImportError:
    from plot_style import apply_publication_style, bolden_axis_text, save_publication_figure
    from survival_60mo import truncate_os_at_months

OS_LANDMARK_MONTHS = 60.0


def _load_result(base: Path, name: str) -> pd.DataFrame:
    return pd.read_csv(base / "result" / name, sep="\t")


def _volcano_on_ax(ax, df: pd.DataFrame, n_labels: int = 5):
    d = df.dropna(subset=["delta_fast_minus_other", "q_diff"]).copy()
    if d.empty:
        ax.set_visible(False)
        return
    d["neglog10q"] = -np.log10(d["q_diff"].clip(lower=1e-300))
    ax.scatter(
        d["delta_fast_minus_other"],
        d["neglog10q"],
        s=16,
        alpha=0.72,
        color="#111827",
        edgecolors="none",
    )
    ax.axvline(0, color="#4b5563", linewidth=1.2)
    ax.set_xlabel("Δβ (short-survival − other)", fontsize=14, fontweight="bold")
    ax.set_ylabel("-log10(q_diff)", fontsize=14, fontweight="bold")
    ax.tick_params(axis="both", which="major", labelsize=12)
    bolden_axis_text(ax)
    lab = d.nsmallest(n_labels, "q_diff")
    offsets = [(6, 8), (8, 24), (-8, -20), (8, 40), (-8, -36)]
    for (_, r), (dx, dy) in zip(lab.iterrows(), offsets):
        ax.annotate(
            str(r["gene_token"]),
            (r["delta_fast_minus_other"], r["neglog10q"]),
            xytext=(dx, dy),
            textcoords="offset points",
            ha="left" if dx > 0 else "right",
            fontsize=10,
            fontweight="bold",
            fontstyle="italic",
            color="#111827",
            alpha=1.0,
            arrowprops={"arrowstyle": "-", "color": "#6b7280", "lw": 0.6},
        )


def _bars_on_ax(ax, df: pd.DataFrame, title: str):
    d = df.dropna(subset=["q_diff", "delta_fast_minus_other"]).sort_values("q_diff").copy()
    if d.empty:
        ax.set_visible(False)
        return
    top_hypo = d[d["delta_fast_minus_other"] < 0].head(10)
    top_hyper = d[d["delta_fast_minus_other"] > 0].head(10)
    top = pd.concat([top_hypo, top_hyper], axis=0)
    if top.empty:
        ax.set_visible(False)
        return
    colors = ["#2563eb" if v < 0 else "#f59e0b" for v in top["delta_fast_minus_other"]]
    ax.barh(top["gene_token"].astype(str), top["delta_fast_minus_other"], color=colors)
    ax.axvline(0, color="black", linewidth=1.2)
    ax.set_xlabel("Δβ (short-survival − other)", fontsize=14, fontweight="bold")
    ax.set_title(title, fontsize=15, fontweight="bold")
    ax.tick_params(axis="x", which="major", labelsize=12)
    ax.tick_params(axis="y", which="major", labelsize=12)
    for lab in ax.get_yticklabels():
        lab.set_fontweight("bold")
        lab.set_fontstyle("italic")


def _make_fastdeath_combined_panels(base: Path):
    """Create component plots and the final 2×2 Figure 3."""
    import matplotlib.pyplot as plt
    import seaborn as sns

    apply_publication_style()
    sns.set_style("whitegrid")
    plot_dir = base / "Plots"

    df_e = _load_result(base, "methylation_fast_death_EAC_ESCA.tsv").replace([np.inf, -np.inf], np.nan)
    df_s = _load_result(base, "methylation_fast_death_ESCC.tsv").replace([np.inf, -np.inf], np.nan)

    # Larger canvas + higher DPI for better readability in Word exports.
    fig, axes = plt.subplots(1, 2, figsize=(16, 6.6), sharey=True)
    _volcano_on_ax(axes[0], df_e)
    _volcano_on_ax(axes[1], df_s)
    fig.tight_layout()
    # 600 dpi to keep labels crisp after Word downsampling.
    fig.savefig(plot_dir / "02_fastdeath_methylation_volcano_panel.png", dpi=600, bbox_inches="tight", facecolor="white", edgecolor="none")
    plt.close(fig)

    fig, axes = plt.subplots(1, 2, figsize=(16, 7.6))
    _bars_on_ax(
        axes[0],
        df_e,
        f"EAC: top hypo/hyper in short survival (OS{int(OS_LANDMARK_MONTHS)})",
    )
    _bars_on_ax(
        axes[1],
        df_s,
        f"ESCC: top hypo/hyper in short survival (OS{int(OS_LANDMARK_MONTHS)})",
    )
    fig.tight_layout()
    fig.savefig(plot_dir / "03_fastdeath_top_hypo_hyper_panel.png", dpi=600, bbox_inches="tight", facecolor="white", edgecolor="none")
    plt.close(fig)

    fig, axes = plt.subplots(
        2,
        2,
        figsize=(16, 13),
        gridspec_kw={"height_ratios": [1.0, 1.15]},
    )
    _volcano_on_ax(axes[0, 0], df_e)
    _volcano_on_ax(axes[0, 1], df_s)
    axes[0, 0].set_title("EAC", fontsize=15, fontweight="bold")
    axes[0, 1].set_title("ESCC", fontsize=15, fontweight="bold")
    _bars_on_ax(axes[1, 0], df_e, "EAC ranked gene-level features")
    _bars_on_ax(axes[1, 1], df_s, "ESCC ranked gene-level features")
    axes[0, 0].text(
        -0.14, 1.10, "A", transform=axes[0, 0].transAxes,
        fontsize=18, fontweight="bold", va="top",
    )
    axes[1, 0].text(
        -0.14, 1.10, "B", transform=axes[1, 0].transAxes,
        fontsize=18, fontweight="bold", va="top",
    )
    fig.tight_layout(h_pad=2.4, w_pad=1.8)
    fig.savefig(
        plot_dir / "03_fastdeath_methylation_combined_A_volcano_B_ranked.png",
        dpi=600,
        bbox_inches="tight",
        facecolor="white",
        edgecolor="none",
    )
    plt.close(fig)


def _esca_km_probe_list(base: Path, top_k: int = 3) -> list[str]:
    tab = _load_result(base, "methylation_fast_death_EAC_ESCA.tsv")
    tab = tab.dropna(subset=["cox_q_meth", "cox_beta_meth"]).sort_values("cox_q_meth").copy()
    if tab.empty:
        return []
    meth = pd.read_csv(base / "data_methylation_hm450.txt", sep="\t", low_memory=False)
    probe_col = meth.columns[0]
    meth = meth.set_index(probe_col)
    meth.index = meth.index.astype(str)
    out = []
    for p in tab["probe"].head(50).astype(str).tolist():
        if p in meth.index:
            out.append(p)
        if len(out) >= top_k:
            break
    return out


def _escc_km_probe_list(base: Path, top_k: int = 3) -> list[str]:
    tab = _load_result(base, "methylation_fast_death_ESCC.tsv")
    tab = tab.dropna(subset=["q_diff"]).sort_values("q_diff").head(50)
    meth = pd.read_csv(base / "data_methylation_hm450.txt", sep="\t", low_memory=False)
    probe_col = meth.columns[0]
    meth = meth.set_index(probe_col)
    meth.index = meth.index.astype(str)
    out = []
    for p in tab["probe"].astype(str).tolist():
        if p in meth.index:
            out.append(p)
        if len(out) >= top_k:
            break
    return out


def _plot_km_on_ax(ax, probe: str, hist_label: str, master: pd.DataFrame, meth: pd.DataFrame):
    from lifelines import KaplanMeierFitter
    from lifelines.statistics import logrank_test

    sub = master.copy()
    sub = sub.dropna(subset=["OS_MONTHS", "OS_EVENT"])
    sub["OS_EVENT"] = sub["OS_EVENT"].astype(int)
    sub["SAMPLE_ID"] = sub["SAMPLE_ID"].astype(str)

    x = meth.loc[probe, sub["SAMPLE_ID"]].astype(float)
    df = sub.set_index("SAMPLE_ID")[["OS_MONTHS", "OS_EVENT"]].join(x.rename("beta"), how="inner").dropna()
    t60, e60 = truncate_os_at_months(df["OS_MONTHS"], df["OS_EVENT"], months=OS_LANDMARK_MONTHS)
    df = df.assign(OS_MONTHS_60=t60.to_numpy(), OS_EVENT_60=e60.to_numpy())

    min_n, min_ev = (40, 8) if "EAC" in hist_label else (30, 6)
    if len(df) < min_n or df["OS_EVENT_60"].sum() < min_ev:
        ax.set_visible(False)
        return

    med = float(df["beta"].median())
    df["group"] = np.where(df["beta"] >= med, "High", "Low")
    g1 = df[df["group"] == "High"]
    g0 = df[df["group"] == "Low"]
    lr = logrank_test(
        g1["OS_MONTHS_60"],
        g0["OS_MONTHS_60"],
        event_observed_A=g1["OS_EVENT_60"],
        event_observed_B=g0["OS_EVENT_60"],
    )

    kmf = KaplanMeierFitter()
    for lab, subg in [("High meth.", g1), ("Low meth.", g0)]:
        kmf.fit(subg["OS_MONTHS_60"], event_observed=subg["OS_EVENT_60"], label=f"{lab} (n={len(subg)})")
        kmf.plot(
            ax=ax,
            ci_show=False,
            show_censors=True,
            censor_styles={"marker": "|", "ms": 7, "mew": 1.4},
            linewidth=2.8,
        )
    ax.set_xlim(0, OS_LANDMARK_MONTHS)
    ax.set_xlabel("Months", fontsize=13, fontweight="bold")
    ax.set_ylabel("Survival probability", fontsize=13, fontweight="bold")
    ax.set_title(
        f"{hist_label}\n$\\it{{{probe}}}$\nlog-rank p={lr.p_value:.3g}",
        fontsize=13,
        fontweight="bold",
    )
    ax.legend(loc="lower left", framealpha=0.95, fontsize=11)
    ax.tick_params(axis="both", which="major", labelsize=12, width=1.2)
    for lab in ax.get_xticklabels() + ax.get_yticklabels():
        lab.set_fontweight("bold")


def _make_combined_km_panel(base: Path, top_k: int = 3):
    """Figure 4: 2×3 Kaplan–Meier panels (ESCA row, ESCC row)."""
    import matplotlib.pyplot as plt
    import seaborn as sns

    apply_publication_style()
    sns.set_style("whitegrid")
    plot_dir = base / "Plots"

    master = _load_result(base, "master_clinical_samples.tsv")
    esca = master[master["HISTOLOGY"] == "EAC/ESCA"].copy()
    escc = master[master["HISTOLOGY"] == "ESCC"].copy()

    meth = pd.read_csv(base / "data_methylation_hm450.txt", sep="\t", low_memory=False)
    probe_col = meth.columns[0]
    meth = meth.set_index(probe_col)
    meth.index = meth.index.astype(str)
    meth.columns = meth.columns.astype(str)

    pe = _esca_km_probe_list(base, top_k=top_k)
    ps = _escc_km_probe_list(base, top_k=top_k)
    if len(pe) < 1 and len(ps) < 1:
        return

    fig, axes = plt.subplots(2, 3, figsize=(14, 8), sharex=True)
    for j, probe in enumerate(pe[:3]):
        _plot_km_on_ax(axes[0, j], probe, "EAC", esca, meth)
    for j in range(len(pe), 3):
        axes[0, j].set_visible(False)

    for j, probe in enumerate(ps[:3]):
        _plot_km_on_ax(axes[1, j], probe, "ESCC", escc, meth)
    for j in range(len(ps), 3):
        axes[1, j].set_visible(False)

    plt.tight_layout()
    save_publication_figure(fig, plot_dir / "04_KM_methylation_60mo_panel.png")
    plt.close(fig)


def main():
    import matplotlib
    import os

    matplotlib.use(os.environ.get("MPLBACKEND", "Agg"))

    base = Path(__file__).resolve().parents[1]
    _make_fastdeath_combined_panels(base)
    _make_combined_km_panel(base, top_k=3)


if __name__ == "__main__":
    main()
