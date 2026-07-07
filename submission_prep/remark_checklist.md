# REMARK checklist — prognostic model reporting

**Study:** Histology-specific penalized Cox models on TCGA ESCA (Firehose Legacy), OS60 landmark.  
**Placement suggestion:** Supplementary Table S11 (or Methods subsection “REMARK compliance”).  
**Legend:** Y = reported; P = partial; N = not applicable to this secondary analysis.

| REMARK item | Topic | Status | Where reported / note |
|-------------|-------|--------|------------------------|
| 1 | Prognostic (not predictive) intent stated | Y | Introduction (prognostic vs predictive); Abstract |
| 2 | Study design described | Y | Methods — TCGA retrospective cohort; secondary analysis |
| 3 | Setting and data source | Y | Methods — cBioPortal Firehose Legacy; GEO modules |
| 4 | Eligibility / inclusion criteria | Y | Methods — primary tumours, concordant multi-omic profiles, EAC vs ESCC |
| 5 | Treatments described | P | TCGA clinical metadata; neoadjuvant context in Introduction (CROSS) only |
| 6 | Outcome definition | Y | Methods — OS60 landmark (truncate 60 mo; late events censored) |
| 7 | Predictor definition | Y | Methods — paired RNA z-scores + HM450 beta (expr::, meth::); ≤60 genes/lineage |
| 8 | Sample size / events | Y | Results — n and OS60 events per histology; Table 1 / S1 |
| 9 | Patient flow diagram | N | Single public cohort; no screening losses beyond omics overlap |
| 10 | Patient characteristics | P | Clinical baseline optional per reviewer; not central to transfer conclusion |
| 11 | Univariable analysis | P | Methylation–OS univariate Cox in EAC (Table 6 / S6); not for full multi-omic model |
| 12 | Multivariable model building | Y | Methods — elastic-net Cox (penalizer 0.1, l1_ratio 0.5) within lineage |
| 13 | Model specification (full list of variables) | Y | Supplementary Cox coefficient tables; feature list in Methods |
| 14 | Handling of missing data | Y | Methods — median imputation of features before fitting |
| 15 | Events per variable | P | ~60 gene pairs → high EPV concern; penalized Cox used; exploratory framing |
| 16 | Model performance measures | Y | Harrell’s C-index within- and cross-lineage (Table 1 / S1; Figure 1) |
| 17 | Internal validation | P | In-sample C-index primary; GDC 2025 5-fold CV as robustness (Results) |
| 18 | External validation | P | GEO modules = feasibility only, **not** HM450 OS replication (Discussion, Limitations) |
| 19 | Model comparison / incremental value | Y | Pooled stratified vs lineage-specific vs cross-apply (Table 1) |
| 20 | Model presentation (equation/scoring) | P | Coefficients in supplementary TSV; no clinical nomogram (exploratory study) |
| 21 | Interpretation of prognostic effect | Y | Discussion — histology-specific structure; transfer failure |
| 22 | Limitations discussed | Y | Limitations section (landmark, events, no methylation external OS, purity) |
| 23 | Generalizability | Y | Discussion/Limitations — TCGA-only discovery; assay-matched validation needed |
| 24 | Conflict of interest | Y | Competing interests statement |
| 25 | Funding | P | Acknowledgements placeholder for grants |

**Reviewer-facing sentence (optional Methods append):**  
Reporting follows REMARK guidance for prognostic model studies (Altman et al., *J Clin Oncol* 2010; items in Supplementary Table S11).
