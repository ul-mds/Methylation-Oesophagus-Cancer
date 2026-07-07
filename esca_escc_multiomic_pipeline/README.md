## ESCA vs ESCC multi-omic pipeline (TCGA ESCA Firehose Legacy)

This folder contains a reproducible analysis pipeline to produce **histology-specific mechanistic results** for ESCA (adenocarcinoma) vs ESCC (squamous) using:

- Clinical OS/DFS (`data_clinical_patient.txt`, `data_clinical_sample.txt`)
- mRNA expression (RSEM or z-scores)
- HM450 methylation (beta values)
- Mutations (`data_mutations.txt`)

### What it runs

1. **Mediation-style decomposition (per histology)**
   - Estimate methylation → expression effect (direction/strength) per gene
   - Estimate expression association with OS in a Cox model controlling for methylation
   - Compare **indirect effect** proxy (product-of-coefficients) between ESCA vs ESCC
   - Output: top histology-specific mediated genes (FDR-controlled), and “strong in ESCA but not ESCC” tables

2. **Histology-specific prognostic models + cross-apply**
   - Train **ESCA-only** model and test in ESCC
   - Train **ESCC-only** model and test in ESCA
   - Train pooled model with **stratified baseline** (histology strata) and/or **histology interactions**
   - Output: concordance indices + calibration summaries, plus feature-level explanations

3. **Mutation × histology interactions + mechanism**
   - Screen mutation×histology Cox interactions for OS
   - For top hits, connect to:
     - mutation → expression / pathway-score shifts (histology-specific)
     - methylation as a moderator of the mutation→expression link
   - Output: ranked interactions with mechanistic evidence tables

### How to run

From the AML workspace root:

```bash
python -m esca_escc_multiomic_pipeline.run_all
```

Outputs will be written to `esca_escc_multiomic_pipeline/out/`.

### Figures + manuscript (study root `Plots/`)

From the **Esophageal Carcinoma (TCGA, Firehose Legacy)** folder (parent of this pipeline):

```bash
export MPLCONFIGDIR=/tmp/mpl
mkdir -p "$MPLCONFIGDIR"
python3 -m esca_escc_multiomic_pipeline.make_methylation_os_plots
python3 -m esca_escc_multiomic_pipeline.make_plots
python3 -m esca_escc_multiomic_pipeline.patch_manuscript_figures
```

Main paper figures: `01`–`06` (combined fast-death panels, combined KM, mutation plots). Supplemental mediation plots: `07`–`08`.

To refresh Results/Abstract/Discussion prose so every figure and table is cited in narrative form:

```bash
python3 -m esca_escc_multiomic_pipeline.enrich_manuscript_narrative
```

To replace the **Results** section with a single flowing scientific story (transitions + figure/table integration):

```bash
python3 -m esca_escc_multiomic_pipeline.rewrite_results_narrative
```

Do **not** run `enrich_manuscript_narrative` after `rewrite_results_narrative`, or it will overwrite the Results text again. Use `enrich_manuscript_narrative` only for Abstract/Introduction/Discussion updates, or run `rewrite_results_narrative` last.

To place **figures inside Results** right after the paragraphs that introduce them (and drop the trailing figure gallery), run once after `rewrite_results_narrative`:

```bash
python3 -m esca_escc_multiomic_pipeline.inline_figures_in_results
```

Re-running is a no-op if Figure 1 is already between the transfer paragraph and Table 1.

To renumber tables (2A–2D → Tables 2–5; Cox → Table 6; mechanistic → Table 7) and fix in-text citations:

```bash
python3 -m esca_escc_multiomic_pipeline.renumber_tables_manuscript
```

