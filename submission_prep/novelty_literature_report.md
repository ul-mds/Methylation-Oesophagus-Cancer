# Novelty & Literature Positioning Report
## Histology-specific prognostic structure in TCGA esophageal carcinoma

**Date:** 1 July 2026
**Basis:** Manuscript results (`result/` tables) + targeted literature review (2013–2026) on esophageal carcinoma epigenetics and prognostic modelling.

---

## 0. Manuscript findings under review

| # | Finding | Key numbers |
|---|---------|-------------|
| F1 | Histology-specific Cox models discriminate well **within** lineage | EAC C-index 0.894; ESCC 0.901 |
| F2 | Models **fail to transfer** across histology | EAC→ESCC 0.469; ESCC→EAC 0.450 (≈ chance) |
| F3 | Pooled model works **only** with histology-stratified baseline | pooled C-index 0.803 |
| F4 | Lineage-specific **fast-death** methylation programs | ESCC: SHISA3 Δβ −0.28 (q≈3.7e−7); EAC: SIGLEC14/EPHB1/ZNF214/RAB11FIP3 |
| F5 | Fast-death definition **robust in ESCC, cut-off-sensitive in EAC** | probe Jaccard 20% vs 30%: ESCC 1.00, EAC 0.311 |
| F6 | Exploratory mutation×histology + module shifts | MACF1→EMT (EAC), etc. (hypothesis-generating) |

---

## 1. Overall novelty verdict

**Moderate–high conceptual novelty; incremental biomarker novelty.**

The **strongest and most defensible novel contribution** is **F2/F3**: an explicit, quantified demonstration that within-histology prognostic multi-omic models collapse to near-chance when applied across the other esophageal histology, and that pooling recovers performance **only** through histology-stratified baseline hazards. This "transportability-failure" framing is **rarely formalized** in the esophageal literature.

The methylation biomarker findings (**F4**) are **partly novel loci in a partly novel framing** (fast-death tail), but individual genes overlap with prior cancer epigenetics — with one **notable directional contradiction** (SHISA3, see §3) that is itself scientifically interesting.

---

## 2. Finding-by-finding literature positioning

### F1 — Within-histology discrimination (C-index ~0.89–0.90)
**Novelty: LOW (expected).**
Many TCGA-based methylation/expression risk models report AUC/C-index in the 0.8–0.9 range within a single esophageal histology:
- EAC methylation-driven risk model (GPBAR1/OLFM4/FOXI2/CASP10), TCGA — Wang et al., *J Oncol / BioMed Res Int* 2021 (PMC8213479).
- EAC 6-gene risk score, AUC 0.837–0.905 (1/3/5-yr) — BMC Cancer 2025 (s12885-025-15090-z).
- ESCC 12-marker diagnostic + 4-marker prognostic panel — *Multi-omic ESCC methylation*, 2022 (PMC8873499).

Your within-lineage performance is **consistent with**, not superior to, prior work. Position as a **sanity-check baseline**, not a headline.

### F2 — Cross-histology transfer failure (C-index ≈ 0.45–0.47)
**Novelty: HIGH (this is the paper's core contribution).**
- The founding TCGA ESCA paper (Nature 2017) recommended treating EAC and ESCC as **distinct entities**, and pan-cancer cell-of-origin work clustered EAC with STAD and ESCC with HNSC (Pan-Squamous) — Hoadley et al.; Clin Epigenetics 2024 (s13148-024-01638-6). **But these establish molecular distinctness, not prognostic-model non-transportability.**
- Virtually all esophageal prognostic-model papers build and validate **within one histology** and never cross-apply a fixed-coefficient model to the other (BMC Cancer 2025; PMC8213479; PMC8873499). The **negative transfer result itself is essentially unreported**.
- The general ML-oncology literature argues the **opposite** default — pan-cancer models often *help* small cohorts via shared structure (PLOS ONE 2025, pone.0341355). Your result is a **counter-example within a specific lineage pair**, which strengthens novelty.
**Claim you can safely make:** first explicit quantification that fixed EAC/ESCC multi-omic Cox models transfer at ≈ chance across esophageal histologies.

### F3 — Pooling recovers performance only with histology-stratified baseline
**Novelty: MODERATE (novel in this domain; known in biostatistics).**
- Stratified/random baseline hazards for pooled Cox with heterogeneous strata is **well established methodologically** (J Appl Stat 2019, 10.1080/02664763.2019.1573216; Biom J 2010).
- IPD meta-analyses of esophageal trials already use **one-step mixed-effect Cox with trial-specific baseline hazards** (JCO/PMC10553121).
- **What's new:** applying this lens to **molecular prognostic scoring** (not treatment-effect estimation) in EAC vs ESCC, and showing pooling without stratification is misleading. Frame as a **practical reporting imperative**, cite the biostatistics precedent honestly.

### F4 — Lineage-specific fast-death methylation loci
**Novelty: MIXED.**
- **Framing (fast-death shortest-survival tail contrast):** conceptually close to established **short-term vs long-term survivor** methylation designs in **other cancers** (glioblastoma STS/LTS — Clin Epigenetics 2013 PMC3592900; extreme-survival GBM 2024 PMC11341684; prostate lethal vs non-lethal — Clin Transl Med 2021 ctm2.1030). **Not previously applied as a within-histology esophageal contrast** → framing is novel *for esophageal cancer*, but not methodologically unprecedented.
- **Individual EAC loci (SIGLEC14, ZNF214, RAB11FIP3, TMEM63C, HCN2, LINGO1):** no established EAC methylation-prognosis reports found → **novel candidate associations** (discovery-grade).
- **EPHB1:** prior esophageal work links **EphB1 protein expression** to *better* survival (semanticscholar 2995...); your finding is **promoter/gene-body hypomethylation in fast-death** (worse outcome). Directionally plausible but note the expression-based prior.

### F4b — SHISA3 hypomethylation in ESCC fast-death
**Novelty: HIGH + a genuine contradiction worth highlighting.**
- Literature consistently reports SHISA3 **promoter HYPERmethylation → silencing → poor prognosis** in **laryngeal SCC** (BioMed Res Int 2017), **nasopharyngeal carcinoma** (Cancer Res 2019, 79:747), lung and colorectal cancer. SHISA3 is framed as a **tumor suppressor silenced by hypermethylation**.
- **Your ESCC fast-death signal is strong HYPOmethylation** (Δβ −0.278, q≈3.7e−7) associated with the **most aggressive** ESCC deaths — the **opposite direction** to the SCC/tumour-suppressor paradigm.
- **This is a citable, non-trivial discrepancy.** Two honest interpretations: (a) probe/region differs from the classic promoter CpG island (gene-body vs promoter); (b) fast-death tail biology differs from tumour-vs-normal contrasts. **Recommend explicitly addressing this in Discussion** — it is a novelty *and* a robustness question a reviewer will raise.

### F5 — Cut-off sensitivity (fast-death definition robustness)
**Novelty: LOW-MODERATE (methodological rigor, not discovery).**
- Rarely reported explicitly in esophageal biomarker papers; adds credibility. Best used as a **rigor/limitations** strength, not a novelty headline.

### F6 — Mutation×histology interactions + module shifts
**Novelty: LOW (exploratory).**
- Interaction p-value lists and module-score shifts are common; your sample size limits power. Keep as **Supplementary/exploratory**; do not claim mechanism.

---

## 3. Direct-comparison table (closest prior work)

| Your finding | Closest prior work | How you differ |
|--------------|--------------------|----------------|
| EAC↔ESCC prognostic non-transfer | TCGA ESCA 2017; pan-squamous clustering (2024) | They show molecular distinctness; you quantify **prognostic-model** collapse (C-index≈0.45) |
| Pooled needs stratified baseline | IPD Cox mixed-effects (JCO 2023; J Appl Stat 2019) | You apply to **molecular risk scores**, EAC vs ESCC |
| Fast-death tail methylation | GBM/prostate STS-vs-LTS methylation | First **within-histology esophageal** application |
| SHISA3 in ESCC | SHISA3 hypermethylation in LSCC/NPC | **Opposite direction** (hypomethylation, fast-death) |
| EAC prognostic methylation genes | GPBAR1/OLFM4/…, IGFBP7 | Different loci (SIGLEC14/ZNF214/RAB11FIP3), fast-death anchored |
| EphB1 & esophageal survival | EphB1 protein → better OS | You report **methylation** in fast-death; reconcile with expression prior |

---

## 4. Novelty scorecard

| Component | Novelty | Publishability strength |
|-----------|---------|-------------------------|
| Cross-histology transfer failure (F2) | ★★★★☆ | **Headline claim** |
| Stratified-baseline pooling caveat (F3) | ★★★☆☆ | Strong supporting |
| Fast-death framing in esophageal (F4) | ★★★☆☆ | Good |
| Novel EAC loci (F4) | ★★★☆☆ | Discovery-grade |
| SHISA3 direction contradiction (F4b) | ★★★★☆ | High-interest, needs Discussion |
| Cut-off robustness (F5) | ★★☆☆☆ | Rigor |
| Mutation×histology (F6) | ★☆☆☆☆ | Exploratory |

---

## 5. Recommended novelty statement (for cover letter / Discussion)

> To our knowledge, this is the first study to explicitly quantify the **non-transportability of fixed multi-omic prognostic models between esophageal adenocarcinoma and squamous cell carcinoma**, showing that within-histology discrimination (C-index ≈ 0.89–0.90) collapses to near-chance across histology (≈ 0.45–0.47) and that pooled scoring is valid only with histology-stratified baseline hazards. We further describe **lineage-resolved fast-death methylation programs**, including a strong ESCC-specific SHISA3 hypomethylation signal that runs **counter to the promoter-hypermethylation, tumour-suppressor paradigm** reported in other squamous carcinomas — a discrepancy that motivates assay-matched, region-resolved external validation.

---

## 6. Reviewer risks tied to novelty (pre-empt these)

1. **"Within-histology C-index is not new."** → Agree; frame as baseline, headline = transfer failure.
2. **"SHISA3 direction contradicts prior SCC data."** → Address explicitly (promoter vs gene-body; tail vs tumour-normal).
3. **"Transfer failure is expected given known EAC/ESCC distinctness."** → Distinctness ≠ quantified prognostic non-transport; no prior paper cross-applies fixed models.
4. **"Novel loci lack external methylation-OS validation."** → Already in Limitations (no assay-matched HM450+OS cohort found).
5. **"Fast-death = relabelled short-term survivor design."** → Cite GBM/prostate precedent; novelty is esophageal + histology-resolved + landmark OS60.

---

## 7. Bottom line

- **Publishable novelty exists**, concentrated in the **transportability-failure** message and the **SHISA3 directional anomaly**.
- Biomarker-panel novelty alone would be incremental (the field is crowded).
- Position the paper as a **cautionary, methodological + epigenetic-discovery** study, not a "new prognostic signature" paper — this matches Cancer Medicine's soundness-over-impact model and reduces overclaiming risk.
