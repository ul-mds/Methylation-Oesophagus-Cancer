"""Final Cancer Medicine compliance corrections with preservation-first backups."""
from __future__ import annotations

import shutil
import re
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image

ROOT = Path("/Users/senol/Desktop/AML/Esophageal Carcinoma (TCGA, Firehose Legacy)")
PKG = ROOT / "Cancer Medicine submission_package"
MANUSCRIPT = PKG / "02_Manuscript.docx"
FIGURES = PKG / "Figures"
SUP = PKG / "Supplementary"
SUP_TABLES = SUP / "Tables"
RESULT = ROOT / "result"
PIPE_OUT = ROOT / "esca_escc_multiomic_pipeline" / "out"
YELLOW = WD_COLOR_INDEX.YELLOW
CLEAR_FILL = PatternFill(fill_type=None)
FINAL_TITLE = (
    "Histology-dependent prognostic model transportability and candidate methylation "
    "patterns in esophageal carcinoma"
)


def backup_package() -> Path:
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    destination = ROOT / f"Cancer Medicine submission_package_before_final_compliance_{stamp}"
    shutil.copytree(PKG, destination)
    return destination


def clear_runs(paragraph: Paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag in {qn("w:r"), qn("w:hyperlink")}:
            paragraph._p.remove(child)


def set_highlighted(paragraph: Paragraph, text: str) -> None:
    clear_runs(paragraph)
    run = paragraph.add_run(text)
    run.font.highlight_color = YELLOW


def find_starts(doc: Document, prefix: str) -> Paragraph:
    for paragraph in doc.paragraphs:
        if paragraph.text.strip().startswith(prefix):
            return paragraph
    raise ValueError(prefix)


def insert_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    element = OxmlElement("w:p")
    paragraph._p.addnext(element)
    new = Paragraph(element, paragraph._parent)
    if text:
        new.add_run(text)
    return new


def gene_symbols() -> set[str]:
    symbols: set[str] = set()
    for path in [
        RESULT / "methylation_fast_death_EAC_ESCA.tsv",
        RESULT / "methylation_fast_death_ESCC.tsv",
        PIPE_OUT / "mediation_proxy_compare_ESCA_vs_ESCC_OS.tsv",
        PIPE_OUT / "interaction_mechanism_mut_to_expr_modules.tsv",
        PIPE_OUT / "mutation_x_histology_interactions_OS.tsv",
    ]:
        frame = pd.read_csv(path, sep="\t", low_memory=False)
        for column in ("gene", "gene_token", "probe"):
            if column in frame.columns:
                symbols.update(
                    value
                    for value in frame[column].dropna().astype(str)
                    if re.fullmatch(r"[A-Za-z0-9.-]+", value)
                )
    return symbols


def italicize_gene_runs(paragraph: Paragraph, symbols: set[str]) -> None:
    if not paragraph.text:
        return
    candidates = sorted(
        {symbol for symbol in symbols if re.search(rf"(?<![A-Za-z0-9]){re.escape(symbol)}(?![A-Za-z0-9])", paragraph.text)},
        key=len,
        reverse=True,
    )
    if not candidates:
        return
    pattern = re.compile(
        r"(?<![A-Za-z0-9])(" + "|".join(re.escape(symbol) for symbol in candidates) + r")(?![A-Za-z0-9])"
    )
    for run in list(paragraph.runs):
        if not pattern.search(run.text):
            continue
        parts = pattern.split(run.text)
        if len(parts) == 1:
            continue
        run.text = parts[0]
        anchor = run._r
        for part in parts[1:]:
            clone = deepcopy(anchor)
            for text_element in clone.findall(".//" + qn("w:t")):
                text_element.getparent().remove(text_element)
            clone_run = Run(clone, paragraph)
            clone_run.text = part
            clone_run.italic = part in symbols
            anchor.addnext(clone)
            anchor = clone


def update_manuscript() -> None:
    doc = Document(str(MANUSCRIPT))
    replacements = {
        "The primary survival endpoint was": (
            "The primary survival endpoint was overall survival truncated at 60 months (OS60). "
            "Follow-up was capped at 60 months, and deaths occurring after 60 months were "
            "administratively censored. This procedure is not a conventional landmark analysis "
            "because it does not condition on being event-free at a prespecified landmark. Separate "
            "elastic-net Cox models (penalizer = 0.1; l1_ratio = 0.5) were fitted in EAC and ESCC. "
            "Predictors comprised paired RNA-expression z-scores and gene-level HM450 methylation beta "
            "values for up to 60 mediation-priority genes per histology after deduplication. Missing "
            "predictor values were median-imputed before fitting. Because feature selection, model "
            "fitting, and within-lineage evaluation used the same samples without nested resampling, "
            "within-lineage C-indices are apparent estimates. Transportability was evaluated by applying "
            "each fitted model, with coefficients fixed, to the opposite histology and calculating "
            "Harrell’s C-index. A pooled comparator used histology-stratified baseline hazards. "
            "Patient-level bootstrap confidence intervals were calculated for the two cross-lineage "
            "evaluations. The proportional-hazards assumption in each histology-specific penalized model "
            "was assessed using rank-transformed Schoenfeld-residual tests for all 120 fitted terms. One "
            "EAC term had a nominal P value of 0.033, but no term departed from proportional hazards after "
            "Benjamini–Hochberg correction (minimum q = 0.999 in both histologies); no ESCC term was "
            "nominally significant. All eligible public cases were included, so no a priori sample-size "
            "calculation was performed. Tests were two-sided; nominal P < 0.05 and FDR q < 0.05 were used "
            "as descriptive thresholds for these exploratory analyses. Reporting followed the REMARK "
            "recommendations [27] (Supplementary Table S12). Exploratory mediation-style "
            "feature-prioritization analyses used full overall survival and are shown in Supplementary "
            "Figs. S1–S2; they were not OS60 validation analyses."
        ),
        "We examined the shortest-survival tail": (
            "We examined the shortest-survival tail among OS60 decedents within each histology. For a "
            "prespecified tail proportion p, the short-survival group contained max[10, ceiling(p × D)] "
            "decedents with the shortest OS60, where D was the number of OS60 deaths in that histology. "
            "The primary proportion was 20%; therefore, the minimum-n rule produced an effective "
            "proportion greater than 20% in ESCC. Gene-level HM450 beta values were compared between "
            "short-survival samples and all other samples of the same histology using two-sided Welch "
            "tests. We report Δβ = mean(β_short-survival) − mean(β_other) and Benjamini–Hochberg FDR "
            "(q_diff). Complete direction-specific ranks are provided in Supplementary Tables S2–S5. "
            "Sensitivity analyses at 10%, 20%, and 30% report patient-membership stability and overlap "
            "among the 200 lowest-q_diff gene-level features (Supplementary Table S10)."
        ),
        "All data and data analysis are schematized": (
            "The analytical workflow is summarized in Fig. 1. We first quantified cross-histology "
            "transportability of prognostic models (Fig. 2, Supplementary Table S1). We then characterized "
            "short-survival methylation patterns within each histology (Fig. 3a–b, Tables 1–2) and linked "
            "selected gene-level features to survival separation (Fig. 4). Finally, we added exploratory "
            "somatic context using mutation-associated module shifts (Fig. 5a, Supplementary Table S7) "
            "and a mutation×histology interaction screen (Fig. 5b, Supplementary Table S11). External "
            "exploratory checks in GEO are summarized in Supplementary Tables S8–S9. A TCGA GDC (2025) "
            "re-harmonization of overlapping TCGA patients provided a processing-package robustness "
            "check for the transferability endpoint, not an independent cohort."
        ),
        "We next characterized methylation correlates": (
            "We next characterized methylation correlates of short survival separately within each "
            "histology, extending prior reviews of esophageal DNA methylation [29]. Within each histology, "
            "gene-level HM450 features were compared between short-survival samples and all other samples, "
            "with Benjamini–Hochberg correction. Figure 3a presents the feature-wide contrasts, and Fig. 3b "
            "summarizes the leading direction-specific features. Complete EAC and ESCC ranks are provided "
            "in Supplementary Tables S2–S5."
        ),
        "We repeated patient classification using": (
            "We repeated patient classification using 10%, 20%, and 30% short-survival-tail proportions "
            "with a minimum short-survival group of 10 patients. Because ESCC had 32 OS60 deaths, the "
            "minimum-n rule selected the same 10 patients at all three proportions; patient and top-200-"
            "feature Jaccard indices were therefore 1.00. EAC had 44 OS60 deaths: the 10% and 20% "
            "definitions each selected 10 patients, whereas the 30% definition selected 14. The EAC "
            "patient-membership Jaccard index was 0.714 and the top-200-feature Jaccard index was 0.311 "
            "for 20% versus 30%. SHISA3 remained the highest-ranked ESCC feature across thresholds. In "
            "contrast, EAC membership and feature rankings were more sensitive to the 30% threshold. "
            "These findings indicate greater threshold sensitivity in EAC but do not establish a broader "
            "biological program (Supplementary Table S10)."
        ),
        "The genome-wide contrasts showed": (
            "The feature-wide contrasts showed different short-survival methylation patterns in EAC and "
            "ESCC. EAC displayed a broad set of negative Δβ values, whereas ESCC showed a more focal signal "
            "led by SHISA3, followed by C7orf13, PRKG1, and SEMA3E (Fig. 3). These observations are "
            "consistent with previously reported methylation–clinical associations in ESCC [30] but remain "
            "exploratory. Table 1 summarizes the leading EAC hypomethylated gene-level features, and Table "
            "2 summarizes the leading ESCC hypo- and hypermethylated gene-level features. Complete "
            "direction-specific results are provided in Supplementary Tables S2–S5."
        ),
        "Table 1 lists the leading": (
            "Table 1 lists the leading hypomethylated gene-level HM450 features in the EAC subset of "
            "TCGA-ESCA from the primary short-survival comparison. Short-survival was defined by the 20% "
            "rule with a minimum of 10 patients. Mean β values are reported for the short-survival and "
            "comparison groups, and negative Δβ values indicate lower methylation in the short-survival "
            "group. Complete EAC hypomethylated and hypermethylated ranks are provided in Supplementary "
            "Tables S2 and S4, respectively."
        ),
        "Table 2 lists the leading": (
            "Table 2 lists the leading ESCC hypo- and hypermethylated gene-level HM450 features from the "
            "primary short-survival comparison. Under the minimum-n rule, the same 10 ESCC patients were "
            "selected at the 10%, 20%, and 30% proportions. SHISA3 remained the leading hypomethylated "
            "feature across thresholds. Complete ESCC hypomethylated and hypermethylated ranks are provided "
            "in Supplementary Tables S3 and S5. In EAC, univariate Cox analyses prioritized CDC14C, VNN3, "
            "and SUGT1 for display in Kaplan–Meier plots (Table 3 and Fig. 4). ESCC Kaplan–Meier panels were "
            "selected from the short-survival differential analysis. Supplementary Table S6 contains the "
            "complete EAC gene-level methylation–OS60 Cox screen."
        ),
        "As exploratory follow-up": (
            "As exploratory follow-up, we identified candidate methylation patterns associated with the "
            "short-survival phenotype in each histology, rather than a shared pan-esophageal program "
            "(Fig. 3, Tables 1–2, Supplementary Tables S2–S5). These observations complement the "
            "transportability analysis; they are not evidence that methylation differences caused the "
            "models to fail. Gene-level features highlighted by Kaplan–Meier separation remain candidates "
            "for assay-matched external validation."
        ),
        "This study has several limitations": (
            "This study has several limitations. First, the cBioPortal Firehose Legacy methylation matrix "
            "contains gene-level HM450 beta summaries indexed by HGNC symbol and does not retain CpG probe "
            "identifiers; the present results therefore cannot localize effects to individual CpG sites. "
            "Second, the GEO analyses were platform-mismatched feasibility checks rather than assay-matched "
            "survival validation. Third, the omics models were not compared with a comprehensive "
            "clinicopathological model containing TNM stage and treatment. Fourth, truncation at 60 months "
            "excludes later events, and the smaller number of ESCC events limits methylation and mutation-"
            "interaction analyses. Finally, tumor purity was not included in the primary analysis, for "
            "example through ABSOLUTE [32] or a methylation-derived estimate. The results should therefore "
            "be interpreted as a TCGA case study of model transportability rather than as a clinically "
            "deployable signature."
        ),
        "Future validation should remap probes": (
            "Future validation should use CpG-level measurements derived from Illumina HM450 IDAT files "
            "and the corresponding manifest, and should evaluate independent EAC and ESCC cohorts using "
            "matched methylation assays and the same OS60 definition. Models should be compared with "
            "clinicopathological predictors, including TNM stage, age, sex, and treatment, using optimism-"
            "corrected performance estimates. Both unrestricted overall survival and OS60 should be "
            "reported, and tumor purity should be considered. Functional studies may investigate the "
            "biological consequences of SHISA3 methylation, but they cannot by themselves establish causality."
        ),
        "Table 1. Top EAC": "Table 1. Leading EAC short-survival hypomethylated gene-level HM450 features",
        "Short-survival denotes the shortest OS60 among EAC": (
            "Short-survival denotes the shortest OS60 among the EAC subset of TCGA-ESCA decedents under "
            "the 20% rule with a minimum group size of 10. Δβ = mean β (short-survival) − mean β (other); "
            "negative values indicate relative hypomethylation in short-survival. q is the Benjamini–"
            "Hochberg FDR. Complete EAC hypomethylated ranks are in Supporting Information file 2 (Table "
            "S2). EAC, esophageal adenocarcinoma; FDR, false discovery rate; HM450, Illumina "
            "HumanMethylation450; OS60, overall survival truncated at 60 months; TCGA-ESCA, The Cancer "
            "Genome Atlas Esophageal Carcinoma study."
        ),
        "Table 2. Top ESCC": "Table 2. Leading ESCC short-survival hypo- and hypermethylated gene-level HM450 features",
        "Short-survival denotes the shortest OS60 among ESCC": (
            "Short-survival denotes the shortest OS60 among ESCC decedents under the 20% rule with a "
            "minimum group size of 10. Δβ = mean β (short-survival) − mean β (other). Negative and positive "
            "values indicate relative hypo- and hypermethylation, respectively. q is the Benjamini–Hochberg "
            "FDR. Complete ESCC ranks are in Supporting Information file 2 (Tables S3 and S5). ESCC, "
            "esophageal squamous cell carcinoma; FDR, false discovery rate; HM450, Illumina "
            "HumanMethylation450; OS60, overall survival truncated at 60 months."
        ),
        "Table 3. Univariate Cox": "Table 3. EAC univariate Cox associations between gene-level HM450 methylation and OS60",
        "Univariate Cox models of continuous HM450": (
            "Univariate Cox models of continuous gene-level HM450 β versus OS60 were fitted separately for "
            "2,000 prespecified EAC features. Cox β is the log-hazard increment per unit increase in β, and "
            "q is the Benjamini–Hochberg FDR across the screened features. Each displayed fit included 89 "
            "patients and 44 OS60 events. This exploratory screen does not define a multivariable signature. "
            "EAC, esophageal adenocarcinoma; FDR, false discovery rate; HM450, Illumina "
            "HumanMethylation450; OS60, overall survival truncated at 60 months."
        ),
        "Figure 1. Analysis pipeline": (
            "Figure 1. Analysis pipeline for the TCGA-ESCA Firehose Legacy cohort. The workflow shows data "
            "sources (a), preprocessing and the OS60 truncated endpoint (b), penalized Cox modeling and "
            "external GEO feasibility analyses (c), short-survival methylation and survival analyses (d), "
            "and exploratory somatic and mediation analyses (e). BH-FDR, Benjamini–Hochberg false discovery "
            "rate; C-index, concordance index; CV, cross-validation; DFS, disease-free survival; EAC, "
            "esophageal adenocarcinoma; EMT, epithelial–mesenchymal transition; ESCC, esophageal squamous "
            "cell carcinoma; GDC, Genomic Data Commons; GEO, Gene Expression Omnibus; HM450, Illumina "
            "HumanMethylation450; IFN-γ, interferon gamma; lncRNA, long non-coding RNA; MAF, mutation "
            "annotation format; MSigDB, Molecular Signatures Database; OS, overall survival; OS60, overall "
            "survival truncated at 60 months; PH, proportional hazards; RNA, ribonucleic acid; TCGA-ESCA, "
            "The Cancer Genome Atlas Esophageal Carcinoma study."
        ),
        "Figure 3. Short-survival": (
            "Figure 3. Short-survival differential methylation by histology. (a) Volcano plots show Δβ "
            "versus −log10(q_diff) for gene-level HM450 features comparing short-survival with all other "
            "samples within EAC and ESCC. (b) Ranked summaries show the leading hypomethylated and "
            "hypermethylated gene-level features in each histology. Δβ = mean β (short-survival) − mean β "
            "(other); q_diff is the Benjamini–Hochberg FDR. Complete direction-specific ranks are provided "
            "in Supplementary Tables S2–S5. EAC, esophageal adenocarcinoma; ESCC, esophageal squamous cell "
            "carcinoma; FDR, false discovery rate; HM450, Illumina HumanMethylation450."
        ),
        "Figure 4. Kaplan–Meier": (
            "Figure 4. Kaplan–Meier OS60 according to median methylation at selected gene-level HM450 "
            "features. Samples were divided into high- and low-methylation groups using the histology-"
            "specific median β value. The upper row shows EAC features prioritized by methylation–OS60 "
            "association; the lower row shows ESCC features prioritized by short-survival differential "
            "methylation. Vertical marks indicate censored observations. EAC, esophageal adenocarcinoma; "
            "ESCC, esophageal squamous cell carcinoma; HM450, Illumina HumanMethylation450; OS60, overall "
            "survival truncated at 60 months."
        ),
        "Figure 5. Exploratory somatic context": (
            "Figure 5. Exploratory somatic context by histology. (a) Heatmap of mutation-associated changes "
            "in expression-derived module scores (Δ = mutant − wild-type), stratified by histology. (b) "
            "Mutation×histology interaction screen on OS60: x-axis, interaction β (log-hazard); y-axis, "
            "−log10 nominal Wald P value. Results are hypothesis-generating and were not significant after "
            "multiplicity correction across the 80 screened genes. EAC, esophageal adenocarcinoma; EMT, "
            "epithelial–mesenchymal transition; ESCC, esophageal squamous cell carcinoma; IFN-γ, interferon "
            "gamma; OS60, overall survival truncated at 60 months."
        ),
        "Supporting Information file 1.": (
            "Supporting Information file 1. File format: PDF (.pdf). Title of data: Supplementary tables "
            "and figures in first-citation order. Description of data: Tables S1–S12 and Figures S1–S2 as "
            "cited in the manuscript. Gene-level methylation tables show ranked features, and the REMARK "
            "checklist is Table S12."
        ),
        "Supplementary Table S2.": "Supplementary Table S2. EAC: ranked hypomethylated gene-level HM450 features in short-survival versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "Supplementary Table S3.": "Supplementary Table S3. ESCC: ranked hypomethylated gene-level HM450 features in short-survival versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "Supplementary Table S4.": "Supplementary Table S4. EAC: ranked hypermethylated gene-level HM450 features in short-survival versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "Supplementary Table S5.": "Supplementary Table S5. ESCC: ranked hypermethylated gene-level HM450 features in short-survival versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "Supplementary Table S6.": "Supplementary Table S6. EAC: complete univariate Cox screen of continuous gene-level HM450 methylation versus OS60 (2,000 prespecified features; Benjamini–Hochberg FDR).",
        "Supplementary Table S10.": (
            "Supplementary Table S10. Sensitivity of the short-survival definition at 10%, 20%, and 30% "
            "in the OS60 cohort. The short-survival group comprised the shortest-survival tail among OS60 "
            "decedents (minimum n = 10). Methylation contrasts compared the short-survival group with all "
            "other samples within each histology using Welch tests and Benjamini–Hochberg FDR. The table "
            "reports patient-membership and top-200-feature Jaccard indices relative to the 20% primary "
            "definition."
        ),
    }
    for prefix, text in replacements.items():
        set_highlighted(find_starts(doc, prefix), text)

    for table in doc.tables:
        header = table.rows[0].cells[0].paragraphs[0]
        if header.text.strip() == "Probe":
            set_highlighted(header, "Gene")

    supp = find_starts(doc, "Supplementary Fig. S1.")
    set_highlighted(
        supp,
        "Supplementary Figure S1. Difference in mediation-style indirect-effect proxies between EAC and "
        "ESCC. The x-axis shows EAC minus ESCC; the y-axis shows −log10 of the smaller expression-"
        "coefficient P value from the two histology-specific models. Full overall survival was used for "
        "exploratory feature prioritization; this is not an OS60 validation analysis. EAC, esophageal "
        "adenocarcinoma; ESCC, esophageal squamous cell carcinoma; OS, overall survival; OS60, overall "
        "survival truncated at 60 months.",
    )
    supp2 = insert_after(
        supp,
        "Supplementary Figure S2. Ranked differences in mediation-style indirect-effect proxies (EAC "
        "minus ESCC) using full overall survival for exploratory feature prioritization. Positive values "
        "are larger in EAC and negative values are larger in ESCC. This is not an OS60 validation analysis "
        "or a clinical signature. EAC, esophageal adenocarcinoma; ESCC, esophageal squamous cell carcinoma; "
        "OS, overall survival; OS60, overall survival truncated at 60 months.",
    )
    for run in supp2.runs:
        run.font.highlight_color = YELLOW

    symbols = gene_symbols()
    for paragraph in doc.paragraphs:
        italicize_gene_runs(paragraph, symbols)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    italicize_gene_runs(paragraph, symbols)

    doc.save(str(MANUSCRIPT))
    clean = Document(str(MANUSCRIPT))
    for highlight in clean._element.findall(".//" + qn("w:highlight")):
        highlight.getparent().remove(highlight)
    clean.save(str(PKG / "02_Manuscript_clean.docx"))


def fitted_size(image_path: Path, max_width: float = 6.4, max_height: float = 6.3) -> tuple[float, float]:
    with Image.open(image_path) as image:
        width, height = image.size
    scale = min(max_width / width, max_height / height)
    return width * scale, height * scale


def replace_embedded_figures(path: Path) -> None:
    doc = Document(str(path))
    for number in range(1, 6):
        legend = find_starts(doc, f"Figure {number}.")
        image_paragraph = Paragraph(legend._p.getnext(), legend._parent)
        if not image_paragraph._p.findall(".//" + qn("a:blip")):
            raise RuntimeError(f"Figure {number} is not immediately after its legend")
        clear_runs(image_paragraph)
        image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        width, height = fitted_size(FIGURES / f"Figure_{number}.png")
        image_paragraph.add_run().add_picture(
            str(FIGURES / f"Figure_{number}.png"),
            width=Inches(width),
            height=Inches(height),
        )
    doc.save(str(path))


def update_cover_letter() -> None:
    path = PKG / "01_Cover_letter_Cancer_Medicine.docx"
    doc = Document(str(path))
    for paragraph in doc.paragraphs:
        if paragraph.text.startswith("cancermed@wiley.com"):
            clear_runs(paragraph)
            paragraph.add_run("cancermed@wiley.com\t8 October 2026")
        elif paragraph.text.startswith("Original Research Article —"):
            clear_runs(paragraph)
            paragraph.add_run(f"Original Research Article — “{FINAL_TITLE}”")
    doc.save(str(path))


def clean_workbook(path: Path) -> None:
    wb = load_workbook(path)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell.value = cell.value.replace("EAC/ESCA", "EAC")
                    cell.value = cell.value.replace("ESCA_only", "EAC_only")
                    if cell.value == "ESCA":
                        cell.value = "EAC"
                cell.fill = CLEAR_FILL
        sheet.sheet_properties.tabColor = None

    if "File_info" in wb.sheetnames:
        info = wb["File_info"]
        for row in range(info.max_row, 0, -1):
            values = " ".join(str(cell.value or "") for cell in info[row])
            if "REVISED FOR REVIEW" in values or "Yellow sheet" in values:
                info.delete_rows(row)

    for name in ("S2", "S3", "S4", "S5", "S6"):
        if name not in wb.sheetnames:
            continue
        sheet = wb[name]
        headers = [cell.value for cell in sheet[1]]
        if "gene_token" in headers:
            sheet.delete_cols(headers.index("gene_token") + 1)
        if sheet["A1"].value == "probe":
            sheet["A1"] = "gene"
        for cell in sheet["A"][1:]:
            cell.font = Font(name=cell.font.name, size=cell.font.sz, bold=cell.font.bold, italic=True)

    if "S10_summary" in wb.sheetnames:
        sheet = wb["S10_summary"]
        renames = {
            "n_probes_q_lt_005": "n_features_q_lt_005",
            "n_top200_probes": "n_top200_features",
            "top200_probe_jaccard": "top200_feature_jaccard",
        }
        for cell in sheet[1]:
            if cell.value in renames:
                cell.value = renames[cell.value]
    if "S10_key_loci" in wb.sheetnames:
        sheet = wb["S10_key_loci"]
        headers = [cell.value for cell in sheet[1]]
        if "probe" in headers:
            sheet.delete_cols(headers.index("probe") + 1)
        renames = {
            "death_tail_pct": "short_survival_tail_pct",
            "genome_wide_q_rank": "feature_wide_q_rank",
        }
        for cell in sheet[1]:
            if cell.value in renames:
                cell.value = renames[cell.value]
    wb.save(path)


def clean_main_table_workbook(path: Path) -> None:
    wb = load_workbook(path)
    sheet = wb.active
    for row in range(sheet.max_row, 0, -1):
        values = " ".join(str(cell.value or "") for cell in sheet[row])
        if "REVISED FOR REVIEW" in values:
            sheet.delete_rows(row)
    for row in sheet.iter_rows():
        for cell in row:
            cell.fill = CLEAR_FILL
    for row in sheet.iter_rows(min_row=1, max_row=min(8, sheet.max_row)):
        if row[0].value in {"probe", "Probe"}:
            row[0].value = "Gene"
            for cell in sheet.iter_rows(min_row=row[0].row + 1, min_col=1, max_col=1):
                cell[0].font = Font(name=cell[0].font.name, size=cell[0].font.sz, italic=True)
            break
    wb.save(path)


def clean_tsvs() -> None:
    s1 = pd.read_csv(SUP_TABLES / "S01_model_performance_cindex.tsv", sep="\t")
    s1["model"] = s1["model"].replace({"ESCA_only": "EAC_only"})
    s1["test"] = s1["test"].replace({"ESCA": "EAC"})
    s1.to_csv(SUP_TABLES / "S01_model_performance_cindex.tsv", sep="\t", index=False)

    for filename in [
        "S02_methylation_fast_death_EAC_ESCA.tsv",
        "S03_methylation_fast_death_ESCC.tsv",
        "S04_methylation_fast_death_EAC_ESCA.tsv",
        "S05_methylation_fast_death_ESCC.tsv",
        "S06_methylation_OS60_Cox_EAC_ESCA.tsv",
    ]:
        path = SUP_TABLES / filename
        frame = pd.read_csv(path, sep="\t")
        frame = frame.drop(columns=["gene_token"], errors="ignore")
        frame = frame.rename(columns={"probe": "gene"})
        frame.to_csv(path, sep="\t", index=False)

    for filename in [
        "S07_interaction_mechanism_mut_to_expr_modules.tsv",
        "S10_summary_fast_death_sensitivity_summary.tsv",
        "S10_key_loci_fast_death_sensitivity_key_loci.tsv",
        "S11_mutation_x_histology_interactions_OS.tsv",
    ]:
        path = SUP_TABLES / filename
        frame = pd.read_csv(path, sep="\t")
        frame = frame.replace({"EAC/ESCA": "EAC"})
        frame = frame.drop(columns=["probe"], errors="ignore")
        frame = frame.rename(
            columns={
                "death_tail_pct": "short_survival_tail_pct",
                "n_probes_q_lt_005": "n_features_q_lt_005",
                "n_top200_probes": "n_top200_features",
                "top200_probe_jaccard": "top200_feature_jaccard",
                "genome_wide_q_rank": "feature_wide_q_rank",
            }
        )
        frame.to_csv(path, sep="\t", index=False)


def update_caption_files() -> None:
    (PKG / "Tables" / "00_Table_captions.txt").write_text(
        "Cancer Medicine — main tables\n\n"
        "Table 1. Leading EAC short-survival hypomethylated gene-level HM450 features.\n"
        "Table 2. Leading ESCC short-survival hypo- and hypermethylated gene-level HM450 features.\n"
        "Table 3. EAC univariate Cox associations between gene-level HM450 methylation and OS60.\n",
        encoding="utf-8",
    )
    (SUP_TABLES / "00_Table_captions.txt").write_text(
        "Cancer Medicine — supporting tables\n\n"
        "Tables S2–S6 contain gene-level HM450 methylation features, not individual CpG probes.\n"
        "Tables S1–S12 are supplied in Supporting_Information_tables.xlsx.\n",
        encoding="utf-8",
    )


def sync_regenerated_figures() -> None:
    plots = ROOT / "Plots"
    copies = {
        plots / "01_cross_apply_cindex.png": FIGURES / "Figure_2.png",
        plots / "03_fastdeath_methylation_combined_A_volcano_B_ranked.png": FIGURES / "Figure_3.png",
        plots / "04_KM_methylation_60mo_panel.png": FIGURES / "Figure_4.png",
        plots / "05_somatic_combined_A_heatmap_B_interaction.png": FIGURES / "Figure_5.png",
        plots / "07_mediation_delta_volcano.png": SUP / "Figures" / "Figure_S1.png",
        plots / "08_top_histology_specific_mediation.png": SUP / "Figures" / "Figure_S2.png",
    }
    for source, destination in copies.items():
        if not source.is_file():
            raise FileNotFoundError(source)
        shutil.copy2(source, destination)


def update_readme() -> None:
    (PKG / "00_README_upload.txt").write_text(
        "Cancer Medicine — upload-ready file guide\n"
        "Official guidelines: https://onlinelibrary.wiley.com/page/journal/20457634/homepage/forauthors.html\n\n"
        "UPLOAD:\n"
        "1. 01_Cover_letter_Cancer_Medicine.docx\n"
        "2. 02_Manuscript_clean.docx (clean editable manuscript; Figures 1–5 are embedded)\n"
        "3. Figures/Figure_1.png through Figures/Figure_5.png (highest-resolution separate files)\n"
        "4. Supplementary/Supporting_Information.pdf\n"
        "5. Supplementary/Supporting_Information_tables.xlsx\n\n"
        "OPTIONAL:\n"
        "- Figures/Figure_1.pdf (vector Figure 1)\n"
        "- Tables/Table_1.xlsx through Table_3.xlsx if the portal requests separate editable tables\n\n"
        "DO NOT UPLOAD:\n"
        "- 02_Manuscript.docx (yellow-highlighted author-review copy)\n"
        "- Figures/Review_highlighted or Supplementary/Review_highlighted\n"
        "- Duplicate numbered aliases, raw TSV working files, caption TXT files, or .DS_Store\n\n"
        "AUTHOR ACTIONS STILL REQUIRED BEFORE SUBMISSION:\n"
        "- Add institutional email addresses for all co-authors to the title page.\n"
        "- Add email addresses for the three suggested reviewers in the cover letter or portal.\n"
        "- Supply the submitting author’s ORCID iD in the portal.\n"
        "- Archive processed data/code in a preservation-grade repository and add its DOI if available.\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    sync_regenerated_figures()
    update_manuscript()
    update_cover_letter()
    clean_tsvs()
    clean_workbook(SUP / "02_Supplementary_Tables.xlsx")
    shutil.copy2(SUP / "02_Supplementary_Tables.xlsx", SUP / "Supporting_Information_tables.xlsx")
    shutil.copy2(
        SUP / "02_Supplementary_Tables.xlsx",
        SUP / "Review_highlighted" / "02_Supplementary_Tables_review.xlsx",
    )
    for name in ("Table_1.xlsx", "Table_2.xlsx", "Table_3.xlsx"):
        clean_main_table_workbook(PKG / "Tables" / name)
    update_caption_files()
    update_readme()
    replace_embedded_figures(MANUSCRIPT)
    replace_embedded_figures(PKG / "02_Manuscript_clean.docx")
    print("Final compliance corrections applied.")
