"""Apply the audited Cancer Medicine revisions and mark review changes in yellow.

The script creates:
- highlighted 02_Manuscript.docx (requested review copy)
- clean 02_Manuscript_clean.docx
- corrected main/supplementary tables and captions
- yellow-annotated review copies of changed figures

It does not invent analyses. Figure and workbook corrections use existing result files.
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageDraw, ImageFont

ROOT = Path("/Users/senol/Desktop/AML/Esophageal Carcinoma (TCGA, Firehose Legacy)")
PKG = ROOT / "Cancer Medicine submission_package"
MANUSCRIPT = PKG / "02_Manuscript.docx"
RESULT = ROOT / "result"
FIGURES = PKG / "Figures"
SUP = PKG / "Supplementary"
SUP_TABLES = SUP / "Tables"
SUP_FIGURES = SUP / "Figures"

YELLOW = WD_COLOR_INDEX.YELLOW
YELLOW_FILL = PatternFill("solid", fgColor="FFF59D")


def clear_runs(paragraph: Paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag in {qn("w:r"), qn("w:hyperlink")}:
            paragraph._p.remove(child)


def set_highlighted(paragraph: Paragraph, text: str) -> None:
    clear_runs(paragraph)
    run = paragraph.add_run(text)
    run.font.highlight_color = YELLOW


def find_exact(doc: Document, text: str) -> Paragraph:
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == text:
            return paragraph
    raise ValueError(f"Paragraph not found: {text!r}")


def find_starts(doc: Document, prefix: str) -> Paragraph:
    for paragraph in doc.paragraphs:
        if paragraph.text.strip().startswith(prefix):
            return paragraph
    raise ValueError(f"Paragraph starting with {prefix!r} not found")


def insert_after(paragraph: Paragraph, text: str, style: str | None = None) -> Paragraph:
    new_el = paragraph._p.makeelement(qn("w:p"), {})
    paragraph._p.addnext(new_el)
    new_p = Paragraph(new_el, paragraph._parent)
    if style:
        new_p.style = style
    run = new_p.add_run(text)
    run.font.highlight_color = YELLOW
    return new_p


def delete_paragraph(paragraph: Paragraph) -> None:
    parent = paragraph._element.getparent()
    if parent is not None:
        parent.remove(paragraph._element)


def highlight_dialect_normalization(doc: Document) -> None:
    """Use consistent US English outside the bibliographic reference list."""
    replacements = {
        "oesophageal": "esophageal",
        "oesophagus": "esophagus",
        "tumour": "tumor",
        "tumours": "tumors",
        "analyse": "analyze",
        "analysed": "analyzed",
        "modelling": "modeling",
        "modelled": "modeled",
        "programmes": "programs",
        "programme": "program",
        "prioritised": "prioritized",
        "prioritisation": "prioritization",
        "re-harmonisation": "re-harmonization",
        "normalised": "normalized",
        "summarise": "summarize",
        "visualized": "visualized",
    }
    in_refs = False
    for paragraph in doc.paragraphs:
        text = paragraph.text
        if text.strip() == "Literature cited":
            in_refs = True
            continue
        if in_refs and text.strip() == "Tables":
            in_refs = False
        if in_refs or not text:
            continue
        revised = text
        for old, new in replacements.items():
            revised = re.sub(rf"\b{re.escape(old)}\b", new, revised, flags=re.IGNORECASE)
        if revised != text:
            set_highlighted(paragraph, revised)


def remove_inline_figures_and_duplicate_captions(doc: Document) -> None:
    """Use separate figure uploads and retain only the final legend section."""
    for paragraph in list(doc.paragraphs):
        if paragraph._element.findall(".//" + qn("w:drawing")):
            delete_paragraph(paragraph)
            continue
        if re.match(r"^Fig\.\s*[1-5]\.", paragraph.text.strip()):
            delete_paragraph(paragraph)


def revise_manuscript() -> None:
    doc = Document(str(MANUSCRIPT))

    # Title page
    set_highlighted(
        doc.paragraphs[0],
        "Histology-specific prognostic architecture in esophageal carcinoma: "
        "cross-histology transport failure and lineage-specific methylation patterns",
    )
    set_highlighted(
        find_starts(doc, "Senol Dogan"),
        "Senol Dogan1,2*, Sanne Hoefnagel3, Chuang Yang4, Rene Thieme5,#, and Toralf Kirsten1,2,#",
    )
    set_highlighted(
        find_starts(doc, "#"),
        "#These authors contributed equally as senior authors.",
    )
    funding = find_starts(doc, "This work was supported by")
    set_highlighted(funding, funding.text.rstrip(".") + ".")
    set_highlighted(
        find_starts(doc, "TCGA esophageal carcinoma (ESCA) Firehose Legacy data"),
        "TCGA Esophageal Carcinoma (ESCA) Firehose Legacy data, including clinical annotations, "
        "HM450 methylation beta values, RNA-expression z-scores, and mutation calls, were obtained "
        "from cBioPortal (https://www.cbioportal.org/study/summary?id=esca_tcga). The GDC "
        "re-harmonization analysis used overlapping TCGA ESCA records from the NCI Genomic Data "
        "Commons (https://portal.gdc.cancer.gov/projects/TCGA-ESCA) and was treated as a robustness "
        "analysis rather than an independent cohort. External expression datasets are available from "
        "GEO under accessions GSE161533, GSE53624, and GSE53622. Processed source tables, figure data, "
        "and analysis scripts are available at https://github.com/SenolDogan/Methylation-Esophagus.",
    )
    set_highlighted(
        find_starts(doc, "S.D. conceived the study"),
        "S.D.: Conceptualization, Methodology, Software, Formal analysis, Visualization, and Writing – "
        "original draft. T.K.: Supervision and Writing – review and editing. R.T., C.Y., and S.H.: "
        "Writing – review and editing. All authors read and approved the final manuscript.",
    )

    # Abstract and keywords
    set_highlighted(
        find_starts(doc, "Background: Esophageal adenocarcinoma"),
        "Background: Esophageal adenocarcinoma (EAC) and esophageal squamous cell carcinoma (ESCC) "
        "are biologically distinct, yet prognostic omics studies often analyze them together. "
        "Methods: Using TCGA ESCA Firehose Legacy data and a 60-month truncated overall survival "
        "endpoint (OS60), we tested whether multi-omic risk scores trained in one histology retained "
        "prognostic discrimination in the other. We also performed exploratory, histology-specific "
        "methylation and somatic analyses. Results: Histology-specific penalized Cox models showed "
        "high apparent, in-sample discrimination (C-index 0.894 for EAC and 0.901 for ESCC). However, "
        "fixed-coefficient cross-application produced near-chance concordance (0.469 and 0.450). A "
        "pooled model with histology-stratified baseline hazards achieved a C-index of 0.803. "
        "Exploratory methylation analyses identified different short-survival-associated patterns in "
        "EAC and ESCC. External GEO analyses assessed workflow portability and expression-level "
        "concordance but did not provide assay-matched validation of the TCGA methylation findings. "
        "Conclusions: Histology should be modeled explicitly in esophageal prognostic studies. "
        "Independent validation with matched molecular assays and aligned clinical endpoints is "
        "required before clinical application.",
    )
    set_highlighted(
        find_starts(doc, "Keywords:"),
        "Keywords: Esophageal carcinoma; DNA methylation; prognostic models; model transportability",
    )

    # Introduction
    set_highlighted(
        find_starts(doc, "If a score is meant to apply"),
        "A prognostic score intended for esophageal carcinoma should retain discrimination in both "
        "EAC and ESCC. Many published models have not been evaluated by applying fixed coefficients "
        "across histologies. Without such testing, transportability remains unknown; lack of testing "
        "should not itself be interpreted as evidence of lineage specificity. This concern is relevant "
        "to recent multi-omics and machine-learning survival models developed within individual cohorts "
        "or histologies [15–20].",
    )

    # Methods
    set_highlighted(
        find_starts(doc, "The analysis sequence is shown in Fig. 1."),
        "The analysis sequence is shown in Fig. 1. Primary tumor data were obtained from the TCGA "
        "Esophageal Carcinoma Firehose Legacy study through cBioPortal [3]. Eligible samples had "
        "concordant clinical data and complete HM450 methylation, RNA-expression, and somatic-mutation "
        "profiles. Histology was assigned from sample-level metadata and OncoTree codes as EAC or ESCC. "
        "Overall survival status and follow-up time were integrated from patient-level records. The "
        "Firehose Legacy dataset was used as the fixed discovery cohort. A separate 2025 GDC "
        "reprocessing of overlapping TCGA patients was used only to assess robustness to data "
        "harmonization and was not considered an independent validation cohort.",
    )
    set_highlighted(
        find_starts(doc, "A 60-month truncated overall survival endpoint"),
        "The primary survival endpoint was overall survival truncated at 60 months (OS60). Follow-up "
        "was capped at 60 months, and deaths occurring after 60 months were administratively censored. "
        "This procedure is not a conventional landmark analysis because it does not condition on being "
        "event-free at a prespecified landmark. Separate elastic-net Cox models (penalizer=0.1; "
        "l1_ratio=0.5) were fitted in EAC and ESCC. Predictors comprised paired RNA-expression z-scores "
        "and HM450 methylation beta values for up to 60 mediation-priority genes per histology after "
        "deduplication. Missing predictor values were median-imputed before fitting. Because feature "
        "selection, model fitting, and within-lineage evaluation used the same samples without nested "
        "resampling, within-lineage C-indices are apparent estimates. Transportability was evaluated by "
        "applying each fitted model, with coefficients fixed, to the opposite histology and calculating "
        "Harrell’s C-index. A pooled comparator used histology-stratified baseline hazards. Patient-level "
        "bootstrap confidence intervals were calculated for the two cross-lineage evaluations and are "
        "reported in the Results. Reporting followed the REMARK recommendations [27] (Supplementary "
        "Table S12). Exploratory mediation-style feature-prioritization analyses used full overall "
        "survival and are shown in Supplementary Figs. S1–S2; they were not OS60 validation analyses.",
    )
    set_highlighted(
        find_starts(doc, "We focused on the short-survival"),
        "We examined the shortest-survival tail among OS60 decedents within each histology. For a "
        "prespecified tail proportion p, the fast-death group contained max[10, ceiling(p × D)] "
        "decedents with the shortest OS60, where D was the number of OS60 deaths in that histology. "
        "The primary proportion was 20%; therefore, the minimum-n rule produced an effective proportion "
        "greater than 20% in ESCC. HM450 beta values were compared between fast-death samples and all "
        "other samples of the same histology using two-sided Welch tests. We report Δβ = mean(β_fast-death) "
        "− mean(β_other) and Benjamini–Hochberg FDR (q_diff). Complete direction-specific ranks are "
        "provided in Supplementary Tables S2–S5. Sensitivity analyses at 10%, 20%, and 30% report "
        "patient-membership stability and overlap among the 200 lowest-q_diff probes (Supplementary "
        "Table S10).",
    )
    set_highlighted(
        find_starts(doc, "External GEO feasibility modules"),
        "For exploratory mutation–histology screening, we considered the 80 most frequently mutated "
        "genes and retained genes with at least eight mutant samples. For each retained gene, we fitted "
        "a Cox model containing mutation status, histology, and their interaction. Nominal interaction "
        "P values were used only to rank candidates and were not interpreted as confirmatory evidence "
        "of effect modification (Supplementary Table S11). We then compared RNA z-scores and expression-"
        "derived module scores between mutant and wild-type samples within each histology using two-sided "
        "Welch tests (Supplementary Table S7). DESeq2 was not used because the input consisted of "
        "normalized RNA z-scores rather than raw sequencing counts. Module scores summarized MSigDB "
        "Hallmark gene sets for epithelial–mesenchymal transition, proliferation, immune response, and "
        "interferon-γ response [28].",
    )
    scope = find_starts(doc, "These GEO modules are feasibility checks")
    set_highlighted(
        scope,
        "The GEO analyses were feasibility checks of biological concordance and workflow portability; "
        "they were not assay-matched validation of TCGA HM450 methylation on survival. GSE161533 "
        "(GPL570) provided paired tumor–normal expression measurements for 28 ESCC pairs. Probes were "
        "mapped to HGNC genes, and the probe with the largest absolute paired change represented each "
        "gene. A union of 110 TCGA ESCC genes was specified before evaluation in GSE161533 "
        "(Supplementary Table S8).",
    )
    p2 = insert_after(
        scope,
        "GSE53624 and GSE53622 used the GPL18109 lncRNA microarray platform and the OS60 endpoint. "
        "The 60 highest-variance lncRNAs in GSE53624 were used to fit an elastic-net Cox model; its "
        "coefficients were held fixed when scoring GSE53622 (Supplementary Table S9). These analyses "
        "tested workflow portability on expression data and did not replicate the HM450 methylation "
        "analysis.",
        "Normal",
    )
    insert_after(
        p2,
        "The 2025 GDC reprocessing used overlapping TCGA patients and therefore served as a robustness "
        "analysis rather than external validation. Five-fold cross-validation reduced the apparent "
        "within-histology C-indices, while cross-histology performance remained poor. A clinical-only "
        "model containing age and sex showed modest discrimination and was treated as a descriptive "
        "baseline rather than a comprehensive clinicopathological comparator.",
        "Normal",
    )

    # Results
    set_highlighted(
        find_starts(doc, "Given the transfer failure above"),
        "We next examined whether methylation correlates of short survival also differed between "
        "histologies, extending prior reviews of esophageal DNA methylation [29]. Within each histology, "
        "HM450 probes were compared between fast-death samples and all other samples, with "
        "Benjamini–Hochberg correction. Figure 3a presents the genome-wide contrasts, and Fig. 3b "
        "summarizes the leading direction-specific loci. Complete EAC and ESCC ranks are provided in "
        "Supplementary Tables S2–S5.",
    )
    set_highlighted(
        find_starts(doc, "The same pattern was recovered in a TCGA GDC"),
        "The GDC reprocessing of overlapping TCGA patients produced the same qualitative pattern. "
        "Apparent within-lineage C-indices were 0.917 in EAC and 0.853 in ESCC, whereas fixed-coefficient "
        "cross-application yielded 0.476 and 0.437. Five-fold cross-validation reduced the mean "
        "within-lineage C-indices to 0.548 in EAC and 0.422 in ESCC, showing that the Firehose "
        "within-lineage estimates were optimistic. In Firehose Legacy, patient-level bootstrap 95% "
        "confidence intervals were 0.354–0.578 for the EAC-trained model applied to ESCC (C-index 0.469) "
        "and 0.360–0.548 for the ESCC-trained model applied to EAC (C-index 0.450). A pooled model with "
        "histology-stratified baseline hazards achieved C-indices of 0.803 in Firehose Legacy and 0.728 "
        "in the GDC robustness analysis. This improvement indicates that histology-related baseline "
        "hazard differences must be modeled explicitly; it does not demonstrate a shared pan-esophageal "
        "molecular signature (Fig. 2 and Supplementary Table S1).",
    )
    for paragraph in list(doc.paragraphs):
        if paragraph.text.startswith("C-index (Harrell’s concordance) on OS60"):
            delete_paragraph(paragraph)

    set_highlighted(
        find_starts(doc, "We repeated patient classification at 10%, 20%, and 30%"),
        "We repeated patient classification using 10%, 20%, and 30% death-tail proportions with a "
        "minimum fast-death group of 10 patients. Because ESCC had 32 OS60 deaths, the minimum-n rule "
        "selected the same 10 patients at all three proportions; patient and top-200-probe Jaccard "
        "indices were therefore 1.00. EAC had 44 OS60 deaths: the 10% and 20% definitions each selected "
        "10 patients, whereas the 30% definition selected 14. The EAC patient-membership Jaccard index "
        "was 0.714 and the top-200-probe Jaccard index was 0.311 for 20% versus 30%. SHISA3 remained the "
        "highest-ranked ESCC locus across thresholds. In contrast, EAC membership and probe rankings "
        "were more sensitive to the 30% threshold. These findings indicate greater threshold sensitivity "
        "in EAC but do not establish a broader biological program (Supplementary Table S10).",
    )
    set_highlighted(
        find_starts(doc, "Genome-wide contrasts revealed divergent fast-death methylation architectures"),
        "The genome-wide contrasts showed different fast-death methylation patterns in EAC and ESCC. "
        "EAC displayed a broad set of negative Δβ values, whereas ESCC showed a more focal signal led "
        "by SHISA3, followed by C7orf13, PRKG1, and SEMA3E (Fig. 3). These observations are consistent "
        "with previously reported methylation–clinical associations in ESCC [30] but remain exploratory. "
        "Table 1 summarizes the leading EAC hypomethylated loci, and Table 2 summarizes the leading ESCC "
        "hypo- and hypermethylated loci. Complete direction-specific results are provided in "
        "Supplementary Tables S2–S5.",
    )
    set_highlighted(
        find_starts(doc, "Table 1 lists ranked HM450 probes"),
        "Table 1 lists the leading EAC/ESCA hypomethylated loci from the primary fast-death comparison. "
        "Fast-death was defined by the 20% rule with a minimum of 10 patients. Mean β values are reported "
        "for the fast-death and comparison groups, and negative Δβ values indicate lower methylation in "
        "the fast-death group. Complete EAC hypomethylated and hypermethylated ranks are provided in "
        "Supplementary Tables S2 and S4, respectively.",
    )
    set_highlighted(
        find_starts(doc, "Table 2 lists ESCC fast-death loci"),
        "Table 2 lists the leading ESCC hypo- and hypermethylated loci from the primary fast-death "
        "comparison. Under the minimum-n rule, the same 10 ESCC patients were selected at the 10%, 20%, "
        "and 30% proportions. SHISA3 remained the leading hypomethylated locus across thresholds. "
        "Complete ESCC hypomethylated and hypermethylated ranks are provided in Supplementary Tables S3 "
        "and S5. In EAC, univariate Cox analyses prioritized CDC14C, VNN3, and GXYLT2 for display in "
        "Kaplan–Meier plots (Table 3 and Fig. 4). ESCC Kaplan–Meier panels were selected from the "
        "fast-death differential analysis. Supplementary Table S6 contains the complete EAC "
        "methylation–OS60 Cox screen.",
    )

    # Discussion
    set_highlighted(
        find_starts(doc, "Molecular distinctness between EAC and ESCC"),
        "The molecular distinction between EAC and ESCC is established, but most esophageal prognostic "
        "models are developed and evaluated within a single histology. Our primary contribution is the "
        "fixed-coefficient cross-application test. Prediction-model guidance distinguishes apparent, "
        "internally validated, and externally validated performance [31]. We therefore interpret the "
        "within-lineage C-indices of approximately 0.89–0.90 as apparent discrimination and as "
        "reproducibility benchmarks, not as estimates of out-of-sample performance. The decline to "
        "approximately 0.45–0.47 after cross-lineage application is the principal finding. Recovery of "
        "pooled discrimination after stratification by histology supports histology-aware modeling but "
        "does not establish a shared molecular signature.",
    )
    set_highlighted(
        find_starts(doc, "Contrasting the shortest-survival tail"),
        "The histology-stratified fast-death analysis identified different discovery-grade methylation "
        "patterns in EAC and ESCC. In ESCC, SHISA3 ranked first by q_diff across the evaluated death-tail "
        "proportions, with Δβ of approximately −0.28 in the primary analysis. Leading EAC candidates "
        "included SIGLEC14, EPHB1, ZNF214, and RAB11FIP3. These loci were identified using exploratory "
        "within-cohort comparisons and should not be interpreted as clinically validated biomarkers. "
        "The methylation findings are consistent with histological heterogeneity but do not demonstrate "
        "that methylation differences caused the failure of cross-histology model transport.",
    )
    set_highlighted(
        find_starts(doc, "Our external GEO modules clarify"),
        "The external GEO analyses clarify what was and was not tested. GSE161533 provided a limited "
        "tumor–normal expression concordance check for a prespecified set of TCGA ESCC genes "
        "(Supplementary Table S8) but did not contain the outcome-linked methylation data required for "
        "validation. GSE53624 and GSE53622 showed that the OS60 penalized-Cox workflow could be applied "
        "to independent lncRNA datasets, but fixed-coefficient transport was weak (Supplementary Table "
        "S9). Neither analysis replicated TCGA HM450 fast-death methylation on survival; assay-matched "
        "methylation validation remains necessary [26,29].",
    )
    set_highlighted(
        find_starts(doc, "The primary source is the cBioPortal Firehose package"),
        "This study has several limitations. First, the Firehose Legacy package provides portal-level "
        "probe-to-gene labels rather than a complete HM450 remapping. Second, the GEO analyses were "
        "platform-mismatched feasibility checks rather than assay-matched survival validation. Third, "
        "the omics models were not compared with a comprehensive clinicopathological model containing "
        "TNM stage and treatment. Fourth, truncation at 60 months excludes later events, and the smaller "
        "number of ESCC events limits methylation and mutation-interaction analyses. Finally, tumor "
        "purity was not included in the primary analysis, for example through ABSOLUTE [32] or a "
        "methylation-derived estimate. The results should therefore be interpreted as a TCGA case study "
        "of model transportability rather than as a clinically deployable signature.",
    )
    set_highlighted(
        find_starts(doc, "These gaps close with a defined follow-up design"),
        "Future validation should remap probes using the Illumina HM450 manifest or GDC IDAT files and "
        "evaluate independent EAC and ESCC cohorts using matched methylation assays and the same OS60 "
        "definition. Models should be compared with clinicopathological predictors, including TNM stage, "
        "age, sex, and treatment, using optimism-corrected performance estimates. Both unrestricted "
        "overall survival and OS60 should be reported, and tumor purity should be considered. Functional "
        "studies may investigate the biological consequences of SHISA3 methylation, but they cannot by "
        "themselves establish causality.",
    )

    # References: insert two validated methodological references and renumber later items.
    references = {
        "27.": (
            "27. McShane, L. M. et al. Reporting recommendations for tumor marker prognostic studies "
            "(REMARK). J. Clin. Oncol. 2005;23:9067–9072. https://doi.org/10.1200/JCO.2004.01.0454."
        ),
        "28.": (
            "28. Liberzon, A. et al. The Molecular Signatures Database (MSigDB) hallmark gene set "
            "collection. Cell Syst. 2015;1:417–425. https://doi.org/10.1016/j.cels.2015.12.004."
        ),
        "29.": (
            "29. Tang, J. et al. DNA methylation in esophageal cancer: technological advances and early "
            "detection clinical applications. Front. Oncol. 2025;15:1543190. "
            "https://doi.org/10.3389/fonc.2025.1543190."
        ),
        "30.": (
            "30. Wang, G. et al. Analysis of the associations between DNA methylation and clinical "
            "features reveals the potential oncogenic effect of CFAP52 in esophageal squamous cell "
            "cancer. Cancer Cell Int. 2025;25:394. https://doi.org/10.1186/s12935-025-03934-4."
        ),
        "31.": (
            "31. Collins, G. S. et al. Transparent reporting of a multivariable prediction model for "
            "individual prognosis or diagnosis (TRIPOD): the TRIPOD statement. BMJ. 2015;350:g7594. "
            "https://doi.org/10.1136/bmj.g7594."
        ),
        "32.": (
            "32. Carter, S. L. et al. Absolute quantification of somatic DNA alterations in human "
            "cancer. Nat. Biotechnol. 2012;30:413–421. https://doi.org/10.1038/nbt.2203."
        ),
    }
    p27 = find_starts(doc, "27. Liberzon")
    p28 = find_starts(doc, "28. Tang")
    p29 = find_starts(doc, "29. Wang")
    p30 = find_starts(doc, "30. Carter")
    set_highlighted(p27, references["27."])
    set_highlighted(p28, references["28."])
    set_highlighted(p29, references["29."])
    set_highlighted(p30, references["30."])
    new31 = insert_after(p30, references["31."], "Normal")
    insert_after(new31, references["32."], "Normal")
    set_highlighted(
        find_starts(doc, "9. Kunene"),
        "9. Kunene, V. et al. Prognostic markers in oesophageal and gastric cancer review. Are they "
        "ready for clinical practice? ESMO Gastrointest. Oncol. 2024;6:100091. "
        "https://doi.org/10.1016/j.esmogo.2024.100091.",
    )
    set_highlighted(
        find_starts(doc, "18. Tang, P."),
        "18. Tang, P. et al. Integrated machine learning developed a prognosis-related gene signature "
        "to predict prognosis in oesophageal squamous cell carcinoma. J. Cell. Mol. Med. "
        "2024;28:e70171. https://doi.org/10.1111/jcmm.70171.",
    )
    set_highlighted(
        find_starts(doc, "25. Zhao"),
        "25. Zhao, D. et al. Advances in multi-omics for esophageal squamous cell carcinoma: diagnostic, "
        "prognostic, and therapeutic perspectives. Protein Cell. 2026;17:491–527. "
        "https://doi.org/10.1093/procel/pwag005.",
    )

    # Main table notes and final legends.
    set_highlighted(
        find_starts(doc, "Fast-death is the shortest-survival tail among EAC/ESCA decedents"),
        "Fast-death denotes the shortest OS60 among EAC/ESCA decedents under the 20% rule with a "
        "minimum group size of 10. Δβ = mean β (fast-death) − mean β (other); negative values indicate "
        "relative hypomethylation in fast-death. q is the Benjamini–Hochberg FDR. Complete EAC "
        "hypomethylated ranks are in Supporting Information file 2 (Table S2).",
    )
    set_highlighted(
        find_starts(doc, "Definitions of fast-death, Δβ and q follow Table 1"),
        "Fast-death denotes the shortest OS60 among ESCC decedents under the 20% rule with a minimum "
        "group size of 10. Δβ = mean β (fast-death) − mean β (other). Negative and positive values "
        "indicate relative hypo- and hypermethylation, respectively. q is the Benjamini–Hochberg FDR. "
        "Complete ESCC ranks are in Supporting Information file 2 (Tables S3 and S5).",
    )
    set_highlighted(
        find_starts(doc, "Figure 1. Analysis pipeline"),
        "Figure 1. Analysis pipeline for the TCGA ESCA Firehose Legacy cohort. The workflow shows data "
        "sources (a), preprocessing and the OS60 truncated endpoint (b), penalized Cox modeling and "
        "external GEO feasibility analyses (c), fast-death methylation and survival analyses (d), and "
        "exploratory somatic and mediation analyses (e). EAC, esophageal adenocarcinoma; ESCC, esophageal "
        "squamous cell carcinoma; GEO, Gene Expression Omnibus; HM450, Illumina HumanMethylation450; "
        "OS60, overall survival truncated at 60 months.",
    )
    set_highlighted(
        find_starts(doc, "Figure 2. Cross-apply performance"),
        "Figure 2. Cross-histology performance of RNA-plus-HM450 penalized Cox models on OS60. Bars show "
        "apparent within-histology performance, fixed-coefficient cross-histology application, and a "
        "pooled model with histology-stratified baseline hazards. The dashed line indicates chance "
        "concordance (C-index=0.5). EAC, esophageal adenocarcinoma; ESCC, esophageal squamous cell "
        "carcinoma; HM450, Illumina HumanMethylation450; OS60, overall survival truncated at 60 months.",
    )
    set_highlighted(
        find_starts(doc, "Figure 3. Fast-death differential methylation"),
        "Figure 3. Fast-death differential methylation by histology. (a) Volcano plots show Δβ versus "
        "−log10(q_diff) for HM450 probes comparing fast-death with all other samples within EAC and ESCC. "
        "(b) Ranked summaries show the leading hypomethylated and hypermethylated loci in each histology. "
        "Δβ = mean β (fast-death) − mean β (other); q_diff is the Benjamini–Hochberg FDR. Complete "
        "direction-specific ranks are provided in Supplementary Tables S2–S5.",
    )
    set_highlighted(
        find_starts(doc, "Figure 4. Kaplan–Meier overall survival"),
        "Figure 4. Kaplan–Meier OS60 according to median methylation at selected HM450 probes. Samples "
        "were divided into high- and low-methylation groups using the histology-specific median β value. "
        "The upper row shows EAC probes prioritized by methylation–OS60 association; the lower row shows "
        "ESCC probes prioritized by fast-death differential methylation. Vertical marks indicate "
        "censored observations. EAC, esophageal adenocarcinoma; ESCC, esophageal squamous cell carcinoma; "
        "HM450, Illumina HumanMethylation450; OS60, overall survival truncated at 60 months.",
    )
    set_highlighted(
        find_starts(doc, "Figure 5. Exploratory somatic context"),
        "Figure 5. Exploratory somatic context by histology. (a) Heatmap of mutation-associated changes "
        "in expression-derived module scores (Δ = mutant − wild-type), stratified by histology. "
        "(b) Mutation×histology interaction screen on OS60: x-axis, interaction β (log-hazard); y-axis, "
        "−log10 nominal Wald P value. Results are hypothesis-generating and were not significant after "
        "genome-wide multiplicity correction. OS60, overall survival truncated at 60 months.",
    )
    set_highlighted(
        find_starts(doc, "Supplementary Table S12."),
        "Supplementary Table S12. REMARK checklist for prognostic model reporting. Reference: McShane "
        "LM et al., J Clin Oncol. 2005;23:9067–9072.",
    )

    remove_inline_figures_and_duplicate_captions(doc)
    highlight_dialect_normalization(doc)

    doc.save(str(MANUSCRIPT))

    clean_path = PKG / "02_Manuscript_clean.docx"
    clean = Document(str(MANUSCRIPT))
    for element in clean._element.findall(".//" + qn("w:highlight")):
        element.getparent().remove(element)
    clean.save(str(clean_path))


def rebuild_supplementary_tables() -> None:
    eac = pd.read_csv(RESULT / "methylation_fast_death_EAC_ESCA.tsv", sep="\t")
    escc = pd.read_csv(RESULT / "methylation_fast_death_ESCC.tsv", sep="\t")
    common = [
        "probe", "gene_token", "mean_beta_fast", "mean_beta_other",
        "delta_fast_minus_other", "p_diff", "q_diff",
    ]
    outputs = {
        "S2": eac[eac["delta_fast_minus_other"] < 0].sort_values("q_diff")[common],
        "S3": escc[escc["delta_fast_minus_other"] < 0].sort_values("q_diff")[common],
        "S4": eac[eac["delta_fast_minus_other"] > 0].sort_values("q_diff")[common],
        "S5": escc[escc["delta_fast_minus_other"] > 0].sort_values("q_diff")[common],
        "S6": eac.dropna(subset=["cox_q_meth"]).sort_values("cox_q_meth")[
            ["probe", "gene_token", "cox_beta_meth", "cox_p_meth", "cox_q_meth", "cox_n", "cox_events"]
        ],
    }
    for key in ("S2", "S3", "S4", "S5"):
        outputs[key] = outputs[key].rename(
            columns={
                "mean_beta_fast": "mean_beta_short_survival",
                "delta_fast_minus_other": "delta_short_survival_minus_other",
            }
        )
    tsv_names = {
        "S2": "S02_methylation_fast_death_EAC_ESCA.tsv",
        "S3": "S03_methylation_fast_death_ESCC.tsv",
        "S4": "S04_methylation_fast_death_EAC_ESCA.tsv",
        "S5": "S05_methylation_fast_death_ESCC.tsv",
        "S6": "S06_methylation_OS60_Cox_EAC_ESCA.tsv",
    }
    for key, frame in outputs.items():
        frame.to_csv(SUP_TABLES / tsv_names[key], sep="\t", index=False)
    stale_s6 = SUP_TABLES / "S06_methylation_fast_death_EAC_ESCA.tsv"
    if stale_s6.exists():
        stale_s6.unlink()

    xlsx = SUP / "02_Supplementary_Tables.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for key, frame in outputs.items():
            frame.to_excel(writer, sheet_name=key, index=False)

    wb = load_workbook(xlsx)
    if "S1" in wb.sheetnames:
        ws = wb["S1"]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == "OS_60mo_landmark":
                    cell.value = "OS60_truncated"
                    cell.fill = YELLOW_FILL
    if "S12" in wb.sheetnames:
        ws = wb["S12"]
        for row in ws.iter_rows(min_row=2):
            if str(row[0].value) == "6":
                row[3].value = "Methods — OS60 truncated endpoint"
                row[3].fill = YELLOW_FILL
            if str(row[0].value) == "11":
                row[3].value = "Table 3 / Supplementary Table S6"
                row[3].fill = YELLOW_FILL
    if "S10_summary" in wb.sheetnames:
        ws = wb["S10_summary"]
        header_updates = {
            "death_tail_pct": "short_survival_tail_pct",
            "n_fast_death": "n_short_survival",
            "max_os_months_in_fast_group": "max_os_months_in_short_survival_group",
        }
        for cell in ws[1]:
            if cell.value in header_updates:
                cell.value = header_updates[cell.value]
                cell.fill = YELLOW_FILL
    for key in ["S2", "S3", "S4", "S5", "S6"]:
        ws = wb[key]
        ws.sheet_properties.tabColor = "FFD700"
        for cell in ws[1]:
            cell.fill = YELLOW_FILL
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    if "File_info" in wb.sheetnames:
        info = wb["File_info"]
    else:
        info = wb.create_sheet("File_info", 0)
    info["A22"] = "REVISED FOR REVIEW"
    info["A22"].fill = YELLOW_FILL
    info["A22"].font = Font(bold=True)
    info["A23"] = (
        "Yellow sheet tabs/header cells identify corrected S2–S6 mappings. "
        "S6 now contains the EAC methylation–OS60 Cox screen."
    )
    info["A23"].fill = YELLOW_FILL
    info["A23"].alignment = Alignment(wrap_text=True)
    wb.save(xlsx)
    shutil.copy2(xlsx, SUP / "Supporting_Information_tables.xlsx")

    # Review workbook keeps the same corrected data and additional yellow marking.
    review_dir = SUP / "Review_highlighted"
    review_dir.mkdir(exist_ok=True)
    shutil.copy2(xlsx, review_dir / "02_Supplementary_Tables_review.xlsx")

    # REMARK TSV mirrors the corrected workbook.
    s12 = pd.read_excel(xlsx, sheet_name="S12")
    s12.to_csv(SUP_TABLES / "S12_REMARK_checklist.tsv", sep="\t", index=False)


def revise_main_table_workbooks() -> None:
    path = PKG / "Tables" / "Table_1.xlsx"
    wb = load_workbook(path)
    ws = wb.active
    ws["A1"] = "Table 1. EAC: leading short-survival-associated hypomethylated HM450 loci."
    ws["A1"].fill = YELLOW_FILL
    ws["A1"].font = Font(bold=True)
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "fast-death" in cell.value.lower():
                cell.value = re.sub("fast-death", "short-survival", cell.value, flags=re.IGNORECASE)
                cell.fill = YELLOW_FILL
    wb.save(path)


def write_caption_files() -> None:
    (FIGURES / "00_Figure_captions.txt").write_text(
        "Cancer Medicine — corrected main figures\n"
        "Figures are supplied separately at high resolution. Yellow-annotated review copies are in "
        "Figures/Review_highlighted and must not be uploaded.\n\n"
        "Figure 1. Analysis pipeline for the TCGA ESCA Firehose Legacy cohort.\n"
        "Figure 2. Cross-histology performance of RNA-plus-HM450 penalized Cox models.\n"
        "Figure 3. Short-survival differential methylation: (a) volcano plots; (b) ranked loci.\n"
        "Figure 4. Kaplan–Meier OS60 with censoring marks.\n"
        "Figure 5. Mutation-associated modules and mutation×histology interactions.\n",
        encoding="utf-8",
    )
    (PKG / "Tables" / "00_Table_captions.txt").write_text(
        "Cancer Medicine — corrected main tables\n\n"
        "Table 1. EAC short-survival-associated HM450 hypomethylated loci.\n"
        "Table 2. ESCC short-survival-associated HM450 hypo- and hypermethylated loci.\n"
        "Table 3. EAC univariate Cox analysis of probe methylation versus OS60.\n",
        encoding="utf-8",
    )
    (SUP_TABLES / "00_Table_captions.txt").write_text(
        "Cancer Medicine — corrected supporting tables\n\n"
        "Table S1. C-index for histology-specific and pooled Cox models.\n"
        "Table S2. EAC short-survival-associated hypomethylated HM450 probes.\n"
        "Table S3. ESCC short-survival-associated hypomethylated HM450 probes.\n"
        "Table S4. EAC short-survival-associated hypermethylated HM450 probes.\n"
        "Table S5. ESCC short-survival-associated hypermethylated HM450 probes.\n"
        "Table S6. EAC univariate methylation–OS60 Cox screen.\n"
        "Table S7. Mutation-associated expression and module shifts.\n"
        "Table S8. GSE161533 overlap with prespecified ESCC genes.\n"
        "Table S9. GSE53624/GSE53622 lncRNA OS60 Cox transport.\n"
        "Table S10. Short-survival threshold sensitivity.\n"
        "Table S11. Mutation-by-histology interaction screen on OS60.\n"
        "Table S12. REMARK checklist (McShane et al., J Clin Oncol. 2005).\n",
        encoding="utf-8",
    )
    (SUP_FIGURES / "00_Figure_captions.txt").write_text(
        "Cancer Medicine — corrected supporting figures\n\n"
        "Figure S1. Exploratory mediation-style methylation–expression delta analysis using full overall "
        "survival; this is not an OS60 validation analysis.\n\n"
        "Figure S2. Exploratory histology-specific mediation contrasts using full overall survival; "
        "this is not an OS60 validation analysis.\n",
        encoding="utf-8",
    )


def add_review_overlay(src: Path, dest: Path, message: str, boxes: list[tuple[int, int, int, int]] | None = None) -> None:
    source = Image.open(src).convert("RGBA")
    width, source_height = source.size
    banner_h = max(44, source_height // 16)
    image = Image.new("RGBA", (width, source_height + banner_h), (255, 255, 255, 255))
    image.paste(source, (0, banner_h))
    layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    height = image.size[1]
    border = max(6, min(width, height) // 120)
    draw.rectangle(
        (border, banner_h + border, width - border, height - border),
        outline=(255, 215, 0, 255),
        width=border,
    )
    draw.rectangle((0, 0, width, banner_h), fill=(255, 245, 157, 235))
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", max(18, banner_h // 3))
    except OSError:
        font = ImageFont.load_default()
    draw.text((16, banner_h // 4), message, fill=(40, 40, 0, 255), font=font)
    for box in boxes or []:
        shifted = (box[0], box[1] + banner_h, box[2], box[3] + banner_h)
        draw.rectangle(shifted, outline=(255, 215, 0, 255), width=border)
    Image.alpha_composite(image, layer).convert("RGB").save(dest, quality=95, dpi=(300, 300))


def create_review_figure_copies() -> None:
    review = FIGURES / "Review_highlighted"
    review.mkdir(exist_ok=True)
    messages = {
        1: "REVISED: OS60 terminology and ESCC sample size",
        2: "REVISED: reader-facing model labels",
        3: "REVISED: missing ranked-loci panel B added",
        4: "REVISED: censoring marks added",
        5: "UNCHANGED IMAGE: final legend corrected in manuscript",
    }
    for number in range(1, 6):
        source = FIGURES / f"Figure_{number}.png"
        if source.exists():
            add_review_overlay(source, review / f"Figure_{number}_review.png", messages[number])

    sup_review = SUP / "Review_highlighted"
    sup_review.mkdir(exist_ok=True)
    for number in (1, 2):
        source = SUP_FIGURES / f"Figure_S{number}.png"
        if source.exists():
            add_review_overlay(
                source,
                sup_review / f"Figure_S{number}_review.png",
                "REVISED DISCLOSURE: exploratory full-OS analysis; not OS60 validation",
            )


def main() -> None:
    revise_manuscript()
    rebuild_supplementary_tables()
    revise_main_table_workbooks()
    write_caption_files()
    print("Revised manuscript and tabular files. Regenerate/copy figures before review overlays.")


if __name__ == "__main__":
    main()
