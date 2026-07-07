from __future__ import annotations

import os
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from .plot_style import apply_publication_style, save_publication_figure
except ImportError:
    from plot_style import apply_publication_style, save_publication_figure


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
    perf2["label"] = perf2["model"] + " → " + perf2["test"]
    perf2 = perf2.sort_values("c_index")
    ax.barh(perf2["label"], perf2["c_index"], color=["#6b7280" if "→ pooled" in x else "#2563eb" for x in perf2["label"]])
    ax.axvline(0.5, color="#b91c1c", linestyle="--", linewidth=1.4, alpha=0.95)
    ax.set_xlabel("C-index")
    ax.set_title("Histology-specific models: within vs cross-apply (60-month landmark OS)")
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
        ax.set_xlabel("Δ indirect (ESCA − ESCC)")
        ax.set_ylabel("-log10(min b_p)")
        ax.set_title("Histology-specific mediation proxy differences (Cox on full OS; not 60-mo)")
        # label a few extremes
        lab = df.reindex(df["delta_indirect_ab"].abs().sort_values(ascending=False).head(8).index)
        for _, r in lab.iterrows():
            ax.text(
                r["delta_indirect_ab"],
                r["neglog10p"],
                str(r["gene"]),
                fontsize=10,
                color="#111827",
                clip_on=True,
            )
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
        ax.set_xlabel("Δ indirect (ESCA − ESCC)")
        ax.set_title("Most histology-specific mediated candidates (full OS Cox proxy)")
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
        ax.set_title("Mutation×histology interaction screen (60-month landmark OS, nominal)")
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


if __name__ == "__main__":
    main()

