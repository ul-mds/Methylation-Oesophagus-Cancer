"""
Apply reviewer-requested revisions to Manuscript.docx.
All new or replaced text is highlighted yellow. Figures, tables, and
unchanged paragraphs are not modified.
"""
from __future__ import annotations

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from esca_escc_multiomic_pipeline.study_paths import study_root

DOC_PATH = study_root() / "Manuscript.docx"


def yellow_run(paragraph: Paragraph, text: str) -> None:
    r = paragraph.add_run(text)
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW


def replace_paragraph_yellow(paragraph: Paragraph, text: str) -> None:
    paragraph.clear()
    yellow_run(paragraph, text)


def insert_paragraph_after(paragraph: Paragraph, text: str) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    yellow_run(new_para, text)
    return new_para


def append_yellow(paragraph: Paragraph, text: str) -> None:
    if paragraph.text and not paragraph.text.endswith((" ", "\t")):
        yellow_run(paragraph, " ")
    yellow_run(paragraph, text)


def find_paragraph(doc: Document, prefix: str) -> Paragraph:
    for p in doc.paragraphs:
        if (p.text or "").startswith(prefix):
            return p
    raise RuntimeError(f"Paragraph not found: {prefix[:80]!r}")


def find_paragraph_contains(doc: Document, needle: str) -> Paragraph:
    for p in doc.paragraphs:
        if needle in (p.text or ""):
            return p
    raise RuntimeError(f"Paragraph not found containing: {needle[:80]!r}")


def main() -> None:
    doc = Document(DOC_PATH)

    # --- References (bottom-up friendly) ---
    ref20 = find_paragraph(doc, "20. Xu, Y.")
    prev = ref20
    new_refs = [
        "21. Spechler SJ, Souza RF. Barrett's esophagus. N Engl J Med. 2014;371:836–845.",
        "22. Thrift AP. The epidemic of oesophageal carcinoma: Where are we? Lancet Gastroenterol Hepatol. 2016;1:608–610.",
        "23. van de Velde CJH, et al. Preoperative chemoradiotherapy for esophageal or junctional cancer. N Engl J Med. 2012;366:2074–2084.",
        "24. Bremner CG, Lynch VP, Ellis FH Jr. Barrett's esophagus: congenital or acquired? Surgery. 1970;68:209–216.",
        "25. Tosh D, Slack JM. How cells change their phenotype. Nat Rev Mol Cell Biol. 2002;3:187–194.",
        "26. Sarosi G, et al. Bone marrow progenitor cells contribute to esophageal regeneration and metaplasia in a rat model of Barrett's esophagus. Dis Esophagus. 2008;21:43–50.",
        "27. Hu Q, et al. Using Barrett's esophagus as a model to study stem cell differentiation. J Gastroenterol. 2014;49:287–294.",
        "28. Mandard AM, et al. Pathologic assessment of tumor regression after preoperative chemoradiotherapy. Int J Radiat Oncol Biol Phys. 1994;29:631–633.",
        "29. Liberzon A, et al. The Molecular Signatures Database (MSigDB) hallmark gene set collection. Cell Syst. 2015;1:417–425.",
        "30. Carter SL, et al. Absolute quantification of somatic DNA alterations in human cancer. Nat Biotechnol. 2012;30:899–907.",
    ]
    prev = ref20
    for ref in new_refs:
        prev = insert_paragraph_after(prev, ref)

    # --- Methods [89] append ---
    append_yellow(
        find_paragraph_contains(doc, "As a robustness check in Esophageal Adenocarcinoma (TCGA GDC, 2025)"),
        "(Primary reporting of this GDC cross-validation appears in Results; clinical baseline comparison is not "
        "central to our cross-histology transfer conclusion.)",
    )

    # --- Methods [88] ---
    replace_paragraph_yellow(
        find_paragraph(doc, "The robustness of our findings was further evaluated through external validation"),
        (
            "Scope. These GEO modules are feasibility checks of biological concordance and workflow portability; they "
            "do not constitute assay-matched external validation of TCGA HM450 fast-death methylation on OS. "
            "GSE161533 (GPL570): paired tumour–normal log2 ratios (+1 pseudocount) across 28 ESCC pairs; probes mapped "
            "to HGNC via mygene.info (max-|Δ| probe per gene). The TCGA gene union (n=110) was pre-specified without "
            "using GSE161533 outcomes: deduplicated union of (i) all protein-coding genes in the ESCC-only penalized Cox "
            "model as expression features (n=60) and (ii) top-80 ESCC mediation-scan genes; overlap with the top 1200 "
            "tumour–normal probes is in Supplementary Table S8. GSE53624/GSE53622 (GPL18109 custom lncRNA microarray): "
            "same OS60 landmark; the 60 highest-variance lncRNAs in GSE53624 only (independent of TCGA methylation/RNA) "
            "trained elastic-net Cox; coefficients were fixed when scoring GSE53622 (Supplementary Table S9)."
        ),
    )

    # --- Methods [87] ---
    replace_paragraph_yellow(
        find_paragraph(doc, "To elucidate the biological underpinnings of lineage-specific survival"),
        (
            "Exploratory mutation–histology interaction screening. On OS60 we screened the 80 most frequently mutated "
            "genes (binary MAF calls), retaining genes with ≥8 mutant samples. For each gene we fitted "
            "hazard ~ mutation + histology + mutation×histology; lineage-dependent survival modulation refers to "
            "nominal mutation×histology interaction p<0.05 (q in Supplementary Table S7). Mechanistic follow-up did "
            "not use DESeq/DESeq2; for retained genes we computed histology-stratified mutant-versus-wild-type contrasts "
            "on RNA z-scores and on module scores (Δ = mean_mutant − mean_wild-type; Welch t-test, nominal p in Table 7). "
            "Module scores were means of within-histology z-normalised expression across MSigDB Hallmark sets: EMT, "
            "G2M Checkpoint+E2F Targets (proliferation), Inflammatory Response+Interferon Alpha Response (immune), and "
            "Interferon Gamma Response (IFN-γ) [29]."
        ),
    )

    # --- Methods [86] heading ---
    replace_paragraph_yellow(
        find_paragraph(doc, "Somatic Contextualization and Multimodal External Validation"),
        "External GEO feasibility modules (expression/lncRNA; not HM450 methylation OS replication)",
    )

    # --- Methods [83] ---
    replace_paragraph_yellow(
        find_paragraph(doc, "Rather than optimizing for broad predictive signatures"),
        (
            'Rather than optimizing for broad prognostic signatures, we focused on "fast-death" phenotypes to identify '
            "methylation programs driving aggressive clinical trajectories. Within each histology, the fast-death group "
            "comprised decedents in the shortest-survival tail (primary definition: top 20% of deaths; minimum n=10). "
            "We compared HM450 beta values between fast-death samples and all other samples of the same histology, "
            "reporting Δβ = mean(β_fast-death) − mean(β_other) with Benjamini–Hochberg FDR (q<0.05). "
            "Sensitivity analyses at 10%, 20%, and 30% death-tail thresholds separately report (i) patient-level "
            "fast-death membership stability and (ii) probe-level overlap of top FDR-significant loci; primary figures "
            "use the pre-specified 20% definition (Results). Under the minimum-n rule, ESCC membership was unchanged "
            "across 10–30% thresholds (32 OS60 deaths), whereas EAC membership varied (44 OS60 deaths; top-200 probe "
            "Jaccard 0.290 for 20% vs 30%)."
        ),
    )

    # --- Methods [80] ---
    replace_paragraph_yellow(
        find_paragraph(doc, "A 60-month clinical landmark (OS60) was established for all survival analyses"),
        (
            "A 60-month clinical landmark (OS60) was established for all survival analyses to align with standard "
            "oncological follow-up windows and minimize bias from late-term attrition. Follow-up time was truncated "
            "at 60 months and events occurring beyond this threshold were treated as censored. To evaluate "
            "inter-histology prognostic transferability, we fitted penalized Cox proportional hazards models "
            "(elastic net; penalizer=0.1, l1_ratio=0.5) separately within EAC and ESCC. Prognostic features comprised "
            "paired RNA expression z-scores and HM450 methylation beta values for the same gene set (expr:: and "
            "meth:: columns), drawn from top mediation-priority genes within each histology (≤60 genes per lineage "
            "after deduplication). Models were evaluated within lineage and by cross-scoring the opposite histology "
            "without refitting coefficients (Harrell's C-index). A pooled comparator used histology-stratified baseline "
            "hazards. Feature-wise median imputation was applied before fitting."
        ),
    )

    # --- Methods [77] ---
    replace_paragraph_yellow(
        find_paragraph(doc, "Primary tumor data were retrieved from the TCGA Esophageal Carcinoma"),
        (
            'Primary tumor data were retrieved from the TCGA Esophageal Carcinoma (ESCA) "Firehose Legacy" dataset '
            "via cBioPortal [1], providing a stable, publicly archived multi-omic snapshot. The inclusion criteria "
            "required samples to have concordant clinical annotations and complete multi-omic profiles, including "
            "Illumina HM450 DNA methylation beta values, RNA expression z-scores, and binary somatic mutation calls. "
            "Histological classification was strictly assigned based on sample-level clinical metadata (OncoTree "
            "codes: EAC vs. ESCC). Clinical endpoints, specifically overall survival (OS) status and time (months), "
            "were integrated from patient-level records. This legacy package was treated as a fixed discovery cohort "
            "for reproducibility; we did not perform additional GDC re-harmonization of the same patient set because "
            "enrolment and follow-up long pre-date the 60-month landmark window."
        ),
    )

    # --- Limitations append ---
    append_yellow(
        find_paragraph(doc, "This study has several limitations that warrant consideration"),
        (
            " Landmark and event-count considerations: the 60-month OS window matches common clinical decision "
            "horizons but limits late events; ESCC contributes fewer OS60 events than EAC, so methylation–OS "
            "screening and mutation×histology interaction tests are exploratory prioritizations rather than "
            "definitive effect-modification evidence. No public series in our search combined HM450 (or successor) "
            "methylation, histology annotation, and harmonised OS suitable to validate the TCGA fast-death "
            "programmes or RNA+HM450 Cox scores out of sample. Tumour purity was not formally estimated (e.g., "
            "ABSOLUTE [30]) in the primary TCGA pipeline."
        ),
    )

    # --- Discussion [65] append ---
    append_yellow(
        find_paragraph(doc, "Finally, we acknowledge the implications of our chosen 60-month landmark horizon"),
        "(Landmark-horizon and event-count caveats are consolidated under Limitations.)",
    )

    # --- Discussion [64] GEO ---
    replace_paragraph_yellow(
        find_paragraph(doc, "Our external validation using GEO datasets provides essential context"),
        (
            "Our external GEO modules clarify what was and was not externally tested. GSE161533 provides limited "
            "tumour–normal expression concordance for TCGA ESCC genes pre-specified before inspecting GSE161533 "
            "(Supplementary Table S8) but lacks outcome-linked validation. GSE53624/GSE53622 (GPL18109 lncRNA "
            "array) confirm that landmark penalized Cox can be executed on independent array data yet show weak "
            "cross-series transport of a fixed lncRNA feature space (Supplementary Table S9). Neither module "
            "replicates TCGA HM450 fast-death methylation on overall survival; assay-matched methylation validation "
            "remains the translational priority [8,20]."
        ),
    )

    # --- Results: fast-death sensitivity insert ---
    insert_paragraph_after(
        find_paragraph(
            doc,
            "Given the transfer failure above, we next tested whether epigenetic correlates",
        ),
        (
            "Sensitivity of the fast-death definition. We repeated patient classification at 10%, 20%, and 30% "
            "death-tail thresholds (minimum n=10 in the fast-death group). Under this rule, ESCC fast-death "
            "membership was unchanged across 10–30% thresholds (32 OS60 deaths), whereas EAC membership varied "
            "more (44 OS60 deaths; top-200 probe overlap Jaccard 0.290 for 20% vs 30%). Leading ESCC loci such "
            "as SHISA3 remained top-ranked across thresholds, supporting a focal ESCC programme; EAC ranks were "
            "more cut-off–sensitive, consistent with a broader adenocarcinoma dispersion."
        ),
    )

    # --- Results: GDC CV append ---
    append_yellow(
        find_paragraph(doc, "This systemic lack of transferability was further substantiated through a parallel analysis of the TCGA GDC"),
        (
            " In the independent TCGA GDC (2025) package, 5-fold cross-validation yielded lower in-lineage "
            "concordance (mean C-index 0.548 in EAC and 0.422 in ESCC) while cross-lineage collapse persisted, "
            "supporting the primary Firehose Legacy transferability finding."
        ),
    )

    # --- Introduction (top-down for inserts) ---
    replace_paragraph_yellow(
        find_paragraph(doc, "In the present study, we prioritize cross-histology transferability"),
        (
            "In the present study, we prioritize cross-histology transferability as a primary evaluative endpoint. "
            "Utilizing penalized Cox proportional hazards models within the TCGA-ESCA framework, we quantified "
            "the internal and cross-lineage performance of prognostic scores over a 60-month overall survival "
            "horizon. Beyond prognostic modeling, we shifted our focus toward an outcome-anchored discovery of "
            "DNA methylation programs. While DNA methylation offers a stable and biologically informative substrate "
            "for biomarker development, many existing panels prioritize prognostic accuracy at the expense of "
            "mechanistic interpretability [5]. By contrasting early attrition cases (rapidly progressing phenotypes) "
            "against other decedents within each lineage, we identified epigenetic programs enriched in aggressive "
            "clinical trajectories."
        ),
    )

    replace_paragraph_yellow(
        find_paragraph(doc, "This methodological oversight creates a significant gap"),
        (
            "This methodological oversight creates a significant gap in the translational utility of public "
            "omics resources. For a prognostic score to be clinically robust as a pan-esophageal tool, it must "
            "maintain its risk-stratification capacity across major histologies. Failure to do so suggests that "
            "the underlying signal is intrinsically tied to the lineage-specific background, meaning that implicit "
            "pooling may obscure rather than clarify biological interpretation. This concern is particularly "
            "relevant for recent multi-omics prognostic constructions and machine-learning frameworks that are "
            "often validated within isolated cohorts or single lineages [12,13,14,15,16,17]."
        ),
    )

    p_intro = find_paragraph(doc, "Esophageal carcinoma is characterized by two predominant histological subtypes")
    replace_paragraph_yellow(
        p_intro,
        (
            "Esophageal carcinoma is characterized by two predominant histological subtypes—adenocarcinoma "
            "(EAC) and squamous cell carcinoma (ESCC)—which exhibit divergent epidemiological profiles and "
            "molecular landscapes. The most important risk factors for ESCC are alcohol use and tobacco "
            "smoking, whereas EAC is associated with the precursor lesion Barrett's esophagus, in which the "
            "normal squamous lining is replaced by abnormal columnar mucosa as a result of "
            "gastro-oesophageal reflux disease. Although both entities arise within the oesophagus, their "
            "anatomic distribution is not identical: ESCC is more commonly located in the mid oesophagus, "
            "whereas EAC typically occurs in the distal oesophagus or gastro-oesophageal junction. Genomic "
            "characterization through The Cancer Genome Atlas (TCGA) has established that these histologies "
            "represent distinct biological entities rather than interchangeable variants [1]. Subsequent "
            "pan-cancer analyses have further delineated lineage-specific transcriptional programs and pathway "
            "architectures [6,7], while reinforcing the association between epigenetic phenotypes and ESCC "
            "subtyping [10,11]. In several European countries, EAC and ESCC are still managed within shared "
            "neoadjuvant chemoradiotherapy frameworks (e.g., the CROSS regimen) [23]. Despite this divergence, "
            "secondary prognostic modeling efforts frequently aggregate EAC and ESCC into a single cohort, and "
            "many lineage-specific signatures are reported without rigorous assessment of cross-histology "
            "prognostic transferability."
        ),
    )

    p_be = insert_paragraph_after(
        p_intro,
        (
            "For EAC, the cellular origin of Barrett's mucosa remains debated. Competing hypotheses include "
            "migration of columnar gastric (stem) cells, trans-differentiation of native oesophageal squamous "
            "cells, contribution of bone-marrow–derived progenitors, and metaplastic expansion from submucosal "
            "gland ducts [24–27]. We therefore use cell-of-origin language when interpreting "
            "histology-resolved epigenetic programmes rather than implying a single uncontested developmental route."
        ),
    )

    insert_paragraph_after(
        p_be,
        (
            "Throughout this manuscript, prognostic denotes association with overall survival (or landmark OS) "
            "independent of a specified therapy [19], whereas predictive is reserved for treatment-response "
            "contexts (e.g., pathological regression after chemoradiotherapy) [28]; we do not use the terms "
            "interchangeably."
        ),
    )

    # --- Abstract ---
    replace_paragraph_yellow(
        find_paragraph(
            doc,
            "Although esophageal adenocarcinoma (EAC) and esophageal squamous cell carcinoma (ESCC) represent pathologically distinct entities, they are frequently aggregated",
        ),
        (
            "Although esophageal adenocarcinoma (EAC) and esophageal squamous cell carcinoma (ESCC) "
            "represent pathologically distinct entities, they are frequently aggregated in secondary "
            "prognostic studies. Utilizing the TCGA ESCA (Firehose Legacy) dataset, we evaluated whether "
            "multi-omic prognostic architectures are transferable across histologies within a standardized "
            "60-month overall survival horizon. Penalized Cox models developed within specific lineages "
            "demonstrated robust internal discrimination (C-index: 0.894 for EAC, 0.901 for ESCC); however, "
            "these models failed to maintain prognostic accuracy when cross-tested, yielding performance "
            "metrics consistent with random assignment (C-index: 0.469 and 0.450). While a pooled modeling "
            "approach recovered cohort-wide performance through histology-stratified baseline hazards "
            "(C-index: 0.803), our findings suggest that pan-esophageal scoring systems require rigorous "
            "cross-lineage validation rather than assumed synergy through pooling. Differential methylation "
            "analysis among decedents revealed lineage-specific epigenetic programs associated with rapid "
            "disease progression. Exploratory mutation-histology interaction screens indicated context-dependent "
            "somatic associations. Modular external GEO checks (GSE161533 tumour–normal expression concordance; "
            "GSE53624→GSE53622 lncRNA prognostic transport) did not replicate HM450 fast-death methylation on "
            "overall survival. These results emphasize histology-aware reporting and assay-matched validation "
            "before omics-based risk stratification is considered clinically actionable."
        ),
    )

    doc.save(DOC_PATH)
    print(f"Saved reviewer revisions (yellow) to:\n  {DOC_PATH}")


if __name__ == "__main__":
    main()
