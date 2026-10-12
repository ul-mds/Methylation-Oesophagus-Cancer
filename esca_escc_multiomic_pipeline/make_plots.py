from __future__ import annotations

import os
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from .plot_style import apply_publication_style, bolden_axis_text, save_publication_figure
except ImportError:
    from plot_style import apply_publication_style, bolden_axis_text, save_publication_figure


def _load(out_dir: Path, name: str) -> pd.DataFrame:
    return pd.read_csv(out_dir / name, sep="\t")


def main():
    import matplotlib

    matplotlib.use(os.environ.get("MPLBACKEND", "Agg"))

    out_dir = Path(__file__).resolve().parent / "out"
    plot_dir = Path(__file__).resolve().parents[1] / "Plots"
    plot_dir.mkdir(parents=True, exist_ok=True)

    import matplotlib.pyplot as plt
    import seaborn as sns

    apply_publication_style()
    sns.set_style("whitegrid")

    # --- Plot 1: Cross-apply performance (C-index) ---
    perf = _load(out_dir, "model_performance_cindex.tsv")
    fig, ax = plt.subplots(figsize=(9, 4))
    perf2 = perf.copy()
    label_map = {
        ("ESCA_only", "ESCA"): "EAC-trained → EAC",
        ("ESCA_only", "ESCC"): "EAC-trained → ESCC",
        ("ESCC_only", "ESCC"): "ESCC-trained → ESCC",
        ("ESCC_only", "ESCA"): "ESCC-trained → EAC",
        ("pooled_strataHIST", "pooled"): "Histology-stratified pooled model",
    }
    perf2["label"] = [
        label_map.get((model, test), f"{model} → {test}")
        for model, test in zip(perf2["model"], perf2["test"])
    ]
    perf2 = perf2.sort_values("c_index")
    ax.barh(
        perf2["label"],
        perf2["c_index"],
        color=["#6b7280" if "pooled" in x.lower() else "#2563eb" for x in perf2["label"]],
    )
    ax.axvline(0.5, color="#b91c1c", linestyle="--", linewidth=1.4, alpha=0.95)
    ax.set_xlabel("C-index", fontweight="bold")
    for lab in ax.get_yticklabels():
        lab.set_fontweight("bold")
    bolden_axis_text(ax)
    save_publication_figure(fig, plot_dir / "01_cross_apply_cindex.png")
    plt.close(fig)

    # --- Plot 2: Mediation compare volcano-like (delta indirect vs min p) ---
    cmp = _load(out_dir, "mediation_proxy_compare_ESCA_vs_ESCC_OS.tsv")
    # pick a conservative significance proxy: min(b_p in either histology)
    cols = ["gene", "delta_indirect_ab", "b_p_ESCA", "b_p_ESCC"]
    if all(c in cmp.columns for c in cols):
        df = cmp[cols].copy()
        df["min_b_p"] = df[["b_p_ESCA", "b_p_ESCC"]].min(axis=1)
        df = df.replace([np.inf, -np.inf], np.nan).dropna(subset=["delta_indirect_ab", "min_b_p"])
        df["neglog10p"] = -np.log10(df["min_b_p"].clip(lower=1e-300))
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.scatter(df["delta_indirect_ab"], df["neglog10p"], s=22, alpha=0.78, color="#111827", edgecolors="none")
        ax.axvline(0, color="#4b5563", linewidth=1.2)
        ax.set_xlabel("Δ indirect effect (EAC − ESCC)")
        ax.set_ylabel("-log10(min b_p)")
        ax.set_title("Histology-specific mediation proxy differences (full OS; not OS60)")
        # label a few extremes
        lab = df.reindex(df["delta_indirect_ab"].abs().sort_values(ascending=False).head(8).index)
        for _, r in lab.iterrows():
            x = float(r["delta_indirect_ab"])
            ax.annotate(
                str(r["gene"]),
                (x, r["neglog10p"]),
                xytext=(-5 if x > 0 else 5, 3),
                textcoords="offset points",
                ha="right" if x > 0 else "left",
                va="bottom",
                fontsize=10,
                fontstyle="italic",
                color="#111827",
                clip_on=False,
            )
        x_pad = 0.08 * (df["delta_indirect_ab"].max() - df["delta_indirect_ab"].min())
        ax.set_xlim(df["delta_indirect_ab"].min() - x_pad, df["delta_indirect_ab"].max() + x_pad)
        save_publication_figure(fig, plot_dir / "07_mediation_delta_volcano.png")
        plt.close(fig)

    # --- Plot 3: Top histology-specific mediated genes (bar) ---
    m = _load(out_dir, "mediation_proxy_compare_ESCA_vs_ESCC_OS.tsv")
    if "delta_indirect_ab" in m.columns and "gene" in m.columns:
        df = m[["gene", "delta_indirect_ab"]].dropna()
        top_pos = df.sort_values("delta_indirect_ab", ascending=False).head(12)
        top_neg = df.sort_values("delta_indirect_ab", ascending=True).head(12)
        top = pd.concat([top_neg, top_pos], axis=0)
        fig, ax = plt.subplots(figsize=(9, 7))
        colors = ["#dc2626" if v < 0 else "#16a34a" for v in top["delta_indirect_ab"]]
        ax.barh(top["gene"], top["delta_indirect_ab"], color=colors)
        ax.axvline(0, color="black", linewidth=1.2)
        ax.set_xlabel("Δ indirect effect (EAC − ESCC)")
        ax.set_title("Most histology-specific mediated candidates (full OS Cox proxy)")
        for label in ax.get_yticklabels():
            label.set_fontstyle("italic")
        save_publication_figure(fig, plot_dir / "08_top_histology_specific_mediation.png")
        plt.close(fig)

    # --- Plot 4: Mutation→module mechanistic examples (heatmap) ---
    mech_path = out_dir / "interaction_mechanism_mut_to_expr_modules.tsv"
    if mech_path.exists():
        mech = _load(out_dir, "interaction_mechanism_mut_to_expr_modules.tsv")
        mech = mech[mech["feature"].astype(str).str.startswith("module::")].copy()
        mech["module"] = mech["feature"].astype(str).str.replace("module::", "", regex=False)
        # keep strongest per gene×histology×module
        mech = mech.sort_values("p").groupby(["gene", "histology", "module"], as_index=False).first()
        # focus on top few genes by best p
        top_genes = mech.sort_values("p").groupby("gene", as_index=False).first().sort_values("p").head(8)["gene"].tolist()
        mech = mech[mech["gene"].isin(top_genes)]
        pivot = mech.pivot_table(index=["gene", "histology"], columns="module", values="delta", aggfunc="first")
        fig, ax = plt.subplots(figsize=(8, 6))
        sns.heatmap(
            pivot,
            ax=ax,
            cmap="vlag",
            center=0,
            linewidths=0.6,
            linecolor="#9ca3af",
            cbar_kws={"shrink": 0.85},
        )
        ax.set_title("Mutation-associated module shifts (Δ=mut−WT; expression, not 60-mo OS)")
        save_publication_figure(fig, plot_dir / "05_mutation_module_heatmap.png")
        plt.close(fig)

    # --- Plot 5: Interaction screen (beta vs -log10p) ---
    inter = _load(out_dir, "mutation_x_histology_interactions_OS.tsv")
    if {"beta_interaction", "p_interaction", "gene"}.issubset(inter.columns):
        df = inter.copy()
        df = df.replace([np.inf, -np.inf], np.nan).dropna(subset=["beta_interaction", "p_interaction"])
        df["neglog10p"] = -np.log10(df["p_interaction"].clip(lower=1e-300))
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.scatter(df["beta_interaction"], df["neglog10p"], s=28, alpha=0.82, color="#1d4ed8", edgecolors="none")
        ax.axvline(0, color="#4b5563", linewidth=1.2)
        ax.set_xlabel("Cox β (mutation×histology interaction)")
        ax.set_ylabel("-log10(p)")
        ax.set_title("Mutation×histology interaction screen (OS60 truncated endpoint, nominal)")
        lab = df.nsmallest(8, "p_interaction")
        for _, r in lab.iterrows():
            ax.text(
                r["beta_interaction"],
                r["neglog10p"],
                str(r["gene"]),
                fontsize=10,
                color="#111827",
                clip_on=True,
            )
        save_publication_figure(fig, plot_dir / "06_interaction_screen.png")
        plt.close(fig)

    # --- Combined Figure 5: heatmap (A) + interaction screen (B) ---
    mech_path = out_dir / "interaction_mechanism_mut_to_expr_modules.tsv"
    inter_path = out_dir / "mutation_x_histology_interactions_OS.tsv"
    if mech_path.exists() and inter_path.exists():
        mech = _load(out_dir, "interaction_mechanism_mut_to_expr_modules.tsv")
        inter = _load(out_dir, "mutation_x_histology_interactions_OS.tsv")
        mech_mod = mech[mech["feature"].astype(str).str.startswith("module::")].copy()
        mech_mod["module"] = mech_mod["feature"].astype(str).str.replace("module::", "", regex=False)
        mech_mod = mech_mod.sort_values("p").groupby(["gene", "histology", "module"], as_index=False).first()
        top_genes = (
            mech_mod.sort_values("p")
            .groupby("gene", as_index=False)
            .first()
            .sort_values("p")
            .head(8)["gene"]
            .tolist()
        )
        mech_sub = mech_mod[mech_mod["gene"].isin(top_genes)].copy()
        mech_sub["gene_hist"] = mech_sub["gene"] + "-" + mech_sub["histology"].astype(str)
        pivot = mech_sub.pivot_table(index="gene_hist", columns="module", values="delta", aggfunc="first")
        module_order = [c for c in ["EMT", "IFNg", "immune", "proliferation"] if c in pivot.columns]
        pivot = pivot.reindex(columns=module_order)

        fig, axes = plt.subplots(2, 1, figsize=(8.5, 10.5), gridspec_kw={"height_ratios": [1.15, 1.0]})
        sns.heatmap(
            pivot,
            ax=axes[0],
            cmap="vlag",
            center=0,
            linewidths=0.6,
            linecolor="#9ca3af",
            cbar_kws={"shrink": 0.85},
        )
        axes[0].set_xlabel("module", fontweight="bold")
        axes[0].set_ylabel("gene-histology", fontweight="bold")
        axes[0].set_xticklabels(axes[0].get_xticklabels(), rotation=0, fontweight="bold")
        gene_hist_labels = []
        for value in pivot.index.astype(str):
            gene, histology = value.split("-", 1)
            histology = "EAC" if histology == "EAC/ESCA" else histology
            gene_hist_labels.append(f"$\\it{{{gene}}}$–{histology}")
        axes[0].set_yticklabels(gene_hist_labels, rotation=0, fontweight="bold")
        module_labels = {
            "EMT": "EMT",
            "IFNg": "IFN-γ",
            "immune": "Immune",
            "proliferation": "Proliferation",
        }
        axes[0].set_xticklabels(
            [module_labels.get(str(label.get_text()), str(label.get_text())) for label in axes[0].get_xticklabels()],
            rotation=0,
            fontweight="bold",
        )
        axes[0].text(-0.10, 1.04, "A", transform=axes[0].transAxes, fontsize=14, fontweight="bold", va="top")

        if {"beta_interaction", "p_interaction", "gene"}.issubset(inter.columns):
            df = inter.replace([np.inf, -np.inf], np.nan).dropna(subset=["beta_interaction", "p_interaction"]).copy()
            df["neglog10p"] = -np.log10(df["p_interaction"].clip(lower=1e-300))
            axes[1].scatter(df["beta_interaction"], df["neglog10p"], s=28, alpha=0.82, color="#1d4ed8", edgecolors="none")
            axes[1].axvline(0, color="#4b5563", linewidth=1.2)
            axes[1].set_xlabel("Cox β (mutation×histology interaction)", fontweight="bold")
            axes[1].set_ylabel("-log10(p)", fontweight="bold")
            lab = df.nsmallest(8, "p_interaction")
            # Focus on the labeled-gene region; extra right padding keeps gene names readable.
            x_lo = min(float(lab["beta_interaction"].min()) - 0.6, -4.5)
            x_hi = max(float(lab["beta_interaction"].max()) + 1.4, 3.6)
            axes[1].set_xlim(x_lo, x_hi)
            for _, r in lab.iterrows():
                bx = float(r["beta_interaction"])
                ha = "right" if bx >= 0 else "left"
                dx = -7 if bx >= 0 else 7
                axes[1].annotate(
                    f"$\\it{{{r['gene']}}}$",
                    (bx, r["neglog10p"]),
                    fontsize=10,
                    fontweight="bold",
                    color="#111827",
                    ha=ha,
                    va="center",
                    xytext=(dx, 0),
                    textcoords="offset points",
                    clip_on=False,
                )
            bolden_axis_text(axes[1])
        axes[1].text(-0.10, 1.04, "B", transform=axes[1].transAxes, fontsize=14, fontweight="bold", va="top")

        fig.tight_layout()
        # Shift panels left within the canvas so panel B gene labels are not clipped on the right.
        pos_a = axes[0].get_position()
        pos_b = axes[1].get_position()
        axes[1].set_position([pos_b.x0 - 0.03, pos_b.y0, pos_b.width * 0.93, pos_b.height])
        axes[0].set_position([pos_a.x0 - 0.03, pos_a.y0, pos_a.width * 0.93, pos_a.height])

        fig.savefig(
            plot_dir / "05_somatic_combined_A_heatmap_B_interaction.png",
            dpi=300,
            bbox_inches="tight",
            pad_inches=0.35,
            facecolor="white",
            edgecolor="none",
        )
        plt.close(fig)


if __name__ == "__main__":
    main()

