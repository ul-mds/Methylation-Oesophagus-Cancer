# Response to Reviewer (internal revision — Sanne Johanna Maria, 26 May 2026)

**Manuscript:** Histology-specific prognostic structure in TCGA esophageal carcinoma  
**Authors:** Senol Dogan, Toralf Kirsten  
**Note:** This letter documents responses to internal Word comments on `Esophagus.docx`. Use as a template if external peer review raises similar points. Comment IDs match the original Word comment numbers.

---

## Comment 0 — Introduction: anatomical origin / Barrett’s cell-of-origin

**Reviewer:** Clarify anatomical overlap vs cell-of-origin; add Barrett’s hypotheses (trans-differentiation, gastric stem cells, bone marrow, submucosal glands) with references.

**Response:** We expanded the Introduction to distinguish anatomical distribution (mid vs distal oesophagus) from debated cell-of-origin mechanisms for Barrett’s mucosa. A new paragraph cites competing hypotheses [24–27] and states that we use cell-of-origin language cautiously when interpreting histology-resolved epigenetic programmes.

**Manuscript location:** Introduction (yellow-highlighted paragraphs); References 24–27.

---

## Comment 41 — Prognostic vs predictive terminology

**Reviewer:** Use “prognostic” only; reserve “predictive” for therapy response (e.g., Mandard).

**Response:** We replaced interchangeable use of “predictive” with “prognostic” throughout (Abstract, Introduction, Methods) and added an explicit definitions paragraph: prognostic = OS association independent of specified therapy [19]; predictive = treatment-response contexts [28].

**Manuscript location:** Abstract; new Introduction paragraph; Methods fast-death section.

---

## Comment 47 — Move 60-month limitations to Limitations

**Reviewer:** Move extended 60-month caveat text to Limitations.

**Response:** The Discussion retains a brief cross-reference; detailed landmark and event-count limitations are consolidated in Limitations (yellow append).

**Manuscript location:** Limitations; Discussion cross-ref sentence.

---

## Comment 54 — GDC re-harmonization expected?

**Reviewer:** Probably not for this legacy cohort enrolled long before the 60-month window.

**Response:** Methods now state that Firehose Legacy is treated as a fixed discovery snapshot and that we did not re-harmonize the same patients via GDC because enrolment predates the landmark window. GDC 2025 is used only as an independent robustness check (Results).

**Manuscript location:** Methods (data source); Results (GDC CV append).

---

## Comment 56 — Cox on RNA expression profiles?

**Reviewer:** Clarify whether models use RNA, methylation, or both.

**Response:** Methods now explicitly state penalized Cox features comprise **paired RNA expression z-scores and HM450 methylation beta values** (expr:: and meth::) for the same gene set.

**Manuscript location:** Methods — Survival Endpoint and Prognostic Modeling.

---

## Comment 61 — Fast-death cut-off sensitivity (10/20/30%): patients vs CpG sites

**Reviewer:** Separate reporting for aggressive **patients** vs aggressive **methylation sites**; unclear which is discussed.

**Response:** We restructured Methods and Results: (1) patient-level fast-death membership stability at 10%, 20%, 30% thresholds; (2) probe-level overlap of top FDR-ranked loci. **Supplementary Table S10** (`result/fast_death_sensitivity_summary.tsv`) quantifies both. ESCC membership is identical across thresholds (min n=10 rule); EAC varies at 30% (n=14).

**Manuscript location:** Methods (fast-death); Results (sensitivity paragraph); Supplementary Table S10.

---

## Comment 70 — External methylation validation?

**Reviewer:** Were methylation profiles validated externally? Any HM450 cohort with OS?

**Response:** We reframed GEO analyses as **feasibility modules**, not assay-matched external validation of TCGA HM450 fast-death methylation on OS. GSE161533 = tumour–normal expression concordance (S8); GSE53624/53622 = lncRNA landmark Cox transport (S9). Limitations state no suitable public HM450+OS+histology series was identified.

**Manuscript location:** Methods heading rewrite; Discussion GEO paragraph; Abstract; Limitations.

---

## Comment 71 — Mutation thresholds and “lineage-dependent survival modulation”

**Reviewer:** Define recurrence threshold; clarify interaction language.

**Response:** Methods specify screening the **80 most frequent mutated genes**, retaining ≥8 mutant samples, and testing mutation×histology interaction on OS60. “Lineage-dependent survival modulation” refers to **nominal mutation×histology interaction p<0.05** (q in S7).

**Manuscript location:** Methods — mutation interaction section.

---

## Comment 72 — DESeq2 / transcriptional comparison groups

**Reviewer:** Which DE method and which patient groups?

**Response:** We did **not** use DESeq/DESeq2. Mechanistic follow-up uses histology-stratified mutant vs wild-type contrasts on RNA z-scores and MSigDB Hallmark module scores (Welch t-test, nominal p in Table 7). Methods now state this explicitly.

**Manuscript location:** Methods — mutation/mechanism paragraph; References [29] MSigDB.

---

## Comment 73 — (empty comment)

**Response:** No text in comment; no action required.

---

## Comment 74 — MSigDB / Broad Institute reference

**Reviewer:** Add reference for module definitions.

**Response:** Hallmark module definitions cite MSigDB [29].

**Manuscript location:** Methods; References 29.

---

## Comment 77 — Pre-specified gene union for GSE161533

**Reviewer:** Pre-specified by whom? Why this gene set?

**Response:** Methods state the 110-gene union was pre-specified **without using GSE161533 outcomes**: all protein-coding genes in the ESCC-only Cox model (n=60) ∪ top-80 ESCC mediation-scan genes.

**Manuscript location:** Methods — GEO GSE161533; Supplementary Table S8.

---

## Comment 88 — Purpose of GSE161533 tumour–normal DE; tumour purity / ABSOLUTE

**Reviewer:** Significance for TCGA analysis? Consider ABSOLUTE purity filtering.

**Response:** GSE161533 tests **expression-level concordance** for TCGA ESCC genes, not methylation OS validation. We did not run ABSOLUTE in the primary pipeline; Limitations note tumour purity was not formally estimated [30] and recommend future purity-aware analyses.

**Manuscript location:** Methods; Discussion; Limitations; Reference 30.

---

## Comment 99 — lncRNA selection linked to TCGA methylation/RNA?

**Reviewer:** Were lncRNAs selected from initial TCGA findings?

**Response:** No. GSE53624/53622 used the **60 highest-variance lncRNAs in GSE53624 only**, independent of TCGA methylation/RNA, to test workflow portability on array data (S9).

**Manuscript location:** Methods — GSE53624/53622.

---

## Comment 113 — Why 60 lncRNAs?

**Reviewer:** Why not all DE lncRNAs?

**Response:** The 60-feature cap matches the approximate feature budget of the TCGA Cox models and limits overfitting in a small training set (n=119 in GSE53624); elastic-net Cox was used. Methods now state this rationale.

**Manuscript location:** Methods — GSE53624/53622.

---

## Comment 116 — GSE53624/53622 platforms; portability scope

**Reviewer:** Clarify platforms; methylation vs other omics portability?

**Response:** Both series use **GPL18109 custom lncRNA microarray** (not HM450). We describe portability of **landmark penalized Cox on lncRNA features**, not methylation risk scores.

**Manuscript location:** Methods — External GEO section.

---

## Comment 117 — GDC cross-validation placement

**Reviewer:** Move GDC CV to Results.

**Response:** GDC 5-fold CV summary moved/appended to Results with cross-lineage collapse note; Methods GDC paragraph retained as robustness with pointer to Results.

**Manuscript location:** Results (yellow append).

---

## Comment 118 — Clinical baseline comparison section

**Reviewer:** Move or omit; omics value already argued in Introduction.

**Response:** Methods GDC paragraph annotated that clinical baseline comparison is not central to cross-histology transfer; full clinical baseline table not required for main conclusion.

**Manuscript location:** Methods (GDC append note).

---

## Summary

All substantive internal comments are addressed in the yellow-highlighted `Manuscript.docx` revisions. Remaining submission tasks: insert **Supplementary Table S10**, **REMARK checklist (S11)**, updated **Data availability** with public code URL, and create GitHub/Zenodo deposit per `submission_prep/CODE_DEPOSIT_README.md`.
