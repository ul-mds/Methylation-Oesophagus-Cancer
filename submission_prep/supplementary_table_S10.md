# Supplementary Table S10. Fast-death death-tail sensitivity (10%, 20%, 30%)

**Caption (for manuscript):** Sensitivity of the fast-death phenotype definition within each histology on the 60-month landmark OS (OS60) cohort. Fast-death membership = shortest-survival tail among OS60 decedents at the indicated death-tail percentage, with minimum n=10 in the fast-death group. Differential methylation contrasts compared HM450 beta values in fast-death samples versus all other samples of the same histology (Welch t-test; Benjamini–Hochberg FDR). Patient membership Jaccard and top-200 probe Jaccard (lowest q_diff ranks) quantify stability between thresholds. Primary analyses use the pre-specified 20% definition. Source: `result/fast_death_sensitivity_summary.tsv` and `result/fast_death_sensitivity_key_loci.tsv`; reproduced by `python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity`.

---

## Panel A. Patient classification and probe-set stability

| Histology | Death-tail (%) | n cohort | n OS60 deaths | n fast-death | Max OS in fast group (mo) | n probes q<0.05 | Patient Jaccard vs prior row* | Top-200 probe Jaccard vs 20%** |
|-----------|----------------|----------|---------------|--------------|---------------------------|-----------------|-------------------------------|--------------------------------|
| EAC/ESCA | 10 | 89 | 44 | 10 | 5.91 | 158 | 1.00 (vs 20%) | 1.00 (vs 20%) |
| EAC/ESCA | 20 | 89 | 44 | 10 | 5.91 | 158 | — (primary) | — (primary) |
| EAC/ESCA | 30 | 89 | 44 | 14 | 7.13 | 62 | 0.714 (vs 20%) | **0.311** (vs 20%) |
| ESCC | 10 | 96 | 32 | 10 | 5.16 | 68 | 1.00 (vs 20%) | 1.00 (vs 20%) |
| ESCC | 20 | 96 | 32 | 10 | 5.16 | 68 | — (primary) | — (primary) |
| ESCC | 30 | 96 | 32 | 10 | 5.16 | 68 | 1.00 (vs 20%) | 1.00 (vs 20%) |

\*Because 10% and 20% yield identical fast-death membership under the minimum-n=10 rule for both histologies, 10% vs 20% Jaccard = 1.00 for EAC and ESCC.

\*\*Top-200 probe Jaccard uses probes with the lowest q_diff at each threshold.

---

## Panel B. Leading loci — genome-wide q_diff rank stability

| Histology | Gene | Rank @ 10% | Rank @ 20% | Rank @ 30% | Stable top rank? |
|-----------|------|------------|------------|------------|------------------|
| ESCC | SHISA3 | 1 | 1 | 1 | Yes |
| ESCC | C7orf13 | 2 | 2 | 2 | Yes |
| ESCC | PRKG1 | 3 | 3 | 3 | Yes |
| ESCC | SEMA3E | 4 | 4 | 4 | Yes |
| EAC/ESCA | SIGLEC14 | 1 | 1 | 176 | No (cut-off sensitive) |
| EAC/ESCA | EPHB1 | 2 | 2 | 81 | Partial |
| EAC/ESCA | ZNF214 | 3 | 3 | 13,533 | No |
| EAC/ESCA | RAB11FIP3 | 5 | 5 | 11 | Partial |

Full numeric detail: `result/fast_death_sensitivity_key_loci.tsv`.

---

## Suggested Results sentence update (replace Jaccard 0.290)

> …whereas EAC membership varied more (44 OS60 deaths; **patient membership Jaccard 0.714 and top-200 probe Jaccard 0.311 for 20% vs 30%**, Supplementary Table S10).

---

## Note on manuscript discrepancy

The draft text used **Jaccard 0.290** for EAC 20% vs 30% probe overlap. The recomputed value from the full HM450 screen is **0.311**. Update Methods/Results to match Supplementary Table S10 before submission.
