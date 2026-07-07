"""
Main-article figure legend bodies (~50 words each including leading 'Figure N. ').
Panel tags use **a**, **b**, **a–c**, etc. for bold formatting in Word.

(Despite the legacy filename, content targets ~50 words, not 300.)
"""
from __future__ import annotations

import re

LEGENDS_300: dict[int, str] = {}

# Idempotency markers for apply_sr_figure_table_guidelines (must appear in each figure body).
LEGEND_IDEMPOTENCY_PHRASE = "histology-specific baseline hazards"
LEGEND2_IDEMPOTENCY_PHRASE = "Hypothesis-generating; replication required"
LEGEND3_IDEMPOTENCY_PHRASE = "Independent histology-specific rankings"
LEGEND4_IDEMPOTENCY_PHRASE = "median split at cohort probe beta"
LEGEND5_IDEMPOTENCY_PHRASE = "Within-histology z-normalised module scores"
LEGEND6_IDEMPOTENCY_PHRASE = "Nominal Wald p-values without genome-wide adjustment"

LEGENDS_300[1] = (
    "RNA-plus-HM450 penalized Cox models on overall survival (OS) test whether prognostic signal transfers between "
    "EAC and ESCC. Harrell’s C-index relates predicted risk rank to survival time ordering (0.5 = chance, 1 = perfect). "
    "**a**, Horizontal bars for within-lineage, cross-lineage, and pooled fits with histology-specific baseline hazards; "
    "ESCA/EAC- versus ESCC-trained models as labelled. Dashed vertical line at 0.5. Point estimates only; Methods give "
    "features and penalties. Pre-specified contrasts only."
)

LEGENDS_300[2] = (
    "Genome-wide HM450 contrasts of fast-death versus other-death among deaths within the analysis window, stratified by histology. "
    "**a**, ESCA/EAC volcano: x-axis, Δβ (mean beta, fast-death minus other deaths); y-axis, −log10 q_diff "
    "(Benjamini–Hochberg FDR from two-group tests on continuous beta). **b**, ESCC, identical scaling; vertical line at "
    "Δβ = 0. Gene symbols highlight extreme probes only; Supplementary Tables give full ranks. Hypothesis-generating; "
    "replication required for mechanism."
)

LEGENDS_300[3] = (
    "Top-ten hypo- and hypermethylated HM450 probes by q_diff for fast-death versus other-death (decedents in the analysis cohort), "
    "matching Figure 2 contrasts. **a**, ESCA/EAC: horizontal bars, length = Δβ on beta scale. **b**, ESCC, same "
    "selection, cool (hypo) and warm (hyper) hues. Each panel is one histology; not formally paired across cohorts. "
    "Gene labels orient probes. "
    + LEGEND3_IDEMPOTENCY_PHRASE
    + "; supplementary tables report means and q_diff."
)

LEGENDS_300[4] = (
    "Kaplan–Meier overall survival by high versus low HM450 methylation ("
    + LEGEND4_IDEMPOTENCY_PHRASE
    + "). Panels Fig. 4a, Fig. 4b, and Fig. 4c, ESCA/EAC probes from Table 6 methylation–OS Cox ranking. Panels Fig. "
    "4d, Fig. 4e, and Fig. 4f, ESCC probes with strong fast-death Δβ (SHISA3, PRKG1, C7orf13). Step functions; tick "
    "marks, censoring; time in months. Log-rank or Cox results in tables."
)

LEGENDS_300[5] = (
    "Exploratory heatmap of mean module score shifts (mutant minus wild-type) for gene-by-histology pairs retained "
    "after filtering. **a**, Rows, gene symbol and lineage; columns, modules with the module:: prefix stripped (Methods). "
    "Colour encodes Δ on a diverging scale centred at zero. "
    + LEGEND5_IDEMPOTENCY_PHRASE
    + ". Cells are cohort means without error bars. Nominal p-values in Table 7; the gene×module grid is not multiplicity-"
    "adjusted. Rare mutations give noisy means."
)

LEGENDS_300[6] = (
    "Exploratory Cox screen for somatic mutation×histology interaction on overall survival. **a**, Each "
    "point is one gene-level model with a mutation×histology product term; x-axis, interaction β (log-hazard); y-axis, "
    "−log10 nominal Wald p; vertical line at β = 0. Semi-transparent markers reduce overlap; gene symbols flag the "
    "smallest nominal p only. "
    + LEGEND6_IDEMPOTENCY_PHRASE
    + ". Quasi-complete separation cases omitted; β and SE in tables (Methods, covariates)."
)


def legend_word_count(n: int) -> int:
    body = LEGENDS_300[n]
    flat = f"Figure {n}. " + re.sub(r"\*\*([a-z](?:–[a-z]+)?)\*\*,\s*", r"\1, ", body)
    return len(flat.split())


def assert_all_near_target() -> None:
    for n in range(1, 7):
        wc = legend_word_count(n)
        lo, hi = (50, 70) if n in (1, 2, 3, 4, 5, 6) else (45, 58)
        if not (lo <= wc <= hi):
            raise ValueError(f"Figure {n} legend word count {wc} outside [{lo},{hi}]")


if __name__ == "__main__":
    assert_all_near_target()
    for n in range(1, 7):
        print(f"Figure {n} legend words (approx): {legend_word_count(n)}")
