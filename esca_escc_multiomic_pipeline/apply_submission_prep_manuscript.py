"""
Apply submission-prep updates to Manuscript.docx (red highlight).
- Fix fast-death Jaccard values (0.290 -> 0.311; add 0.714)
- Supplementary Table S10 (fast-death sensitivity); renumber mutation S10 -> S11
- Supplementary Table S12 (REMARK checklist)
- Data availability + Code availability (placeholders for user)
"""
from __future__ import annotations

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from esca_escc_multiomic_pipeline.study_paths import study_root

DOC_PATH = study_root() / "Manuscript.docx"

RESULTS_SENSITIVITY = (
    "Sensitivity of the fast-death definition. We repeated patient classification at "
    "10%, 20%, and 30% death-tail thresholds (minimum n=10 in the fast-death group). "
    "Under this rule, ESCC fast-death membership was unchanged across 10–30% thresholds "
    "(32 OS60 deaths; n=10 at each threshold; patient and top-200 probe Jaccard 1.00 for "
    "all pairwise comparisons; Supplementary Table S10), whereas EAC membership varied at "
    "30% (44 OS60 deaths; n=10 at 10–20%, n=14 at 30%; patient membership Jaccard 0.714 "
    "and top-200 probe Jaccard 0.311 for 20% vs 30%; Supplementary Table S10). Leading "
    "ESCC loci such as SHISA3 remained genome-wide q_diff rank 1 across thresholds "
    "(Supplementary Table S10), supporting a focal ESCC programme; EAC ranks including "
    "SIGLEC14 were more cut-off–sensitive (Supplementary Table S10), consistent with a "
    "broader adenocarcinoma dispersion."
)

METHODS_FAST_DEATH = (
    'Rather than optimizing for broad prognostic signatures, we focused on "fast-death" '
    "phenotypes to identify methylation programs driving aggressive clinical trajectories. "
    "Within each histology, the fast-death group comprised decedents in the shortest-survival "
    "tail (primary definition: top 20% of deaths; minimum n=10). We compared HM450 beta "
    "values between fast-death samples and all other samples of the same histology, reporting "
    "Δβ = mean(β_fast-death) − mean(β_other) with Benjamini–Hochberg FDR (q<0.05). "
    "Sensitivity analyses at 10%, 20%, and 30% death-tail thresholds separately report "
    "(i) patient-level fast-death membership stability and (ii) probe-level overlap of the "
    "top 200 q_diff-ranked loci (Supplementary Table S10); primary figures use the "
    "pre-specified 20% definition (Results). Under the minimum-n rule, ESCC membership was "
    "unchanged across 10–30% thresholds (32 OS60 deaths), whereas EAC membership varied at "
    "30% (44 OS60 deaths; patient membership Jaccard 0.714 and top-200 probe Jaccard 0.311 "
    "for 20% vs 30%; Supplementary Table S10)."
)

DATA_AVAILABILITY = (
    "TCGA esophageal carcinoma (ESCA) Firehose Legacy data (clinical annotations, HM450 "
    "methylation beta values, RNA expression z-scores, and mutation calls) were accessed "
    "through cBioPortal (https://www.cbioportal.org/study/summary?id=esca_tcga). External "
    "GEO series used for modular feasibility analyses were GSE161533 (GPL570; paired "
    "tumour–normal expression), GSE53624 (GPL18109; lncRNA microarray, training), and "
    "GSE53622 (GPL18109; lncRNA microarray, hold-out scoring). Processed tabular outputs "
    "supporting all figures and tables (including Supplementary Tables S1–S12) are in the "
    "project result/ directory; figure images are in Plots/. Analysis code to reproduce the "
    "primary TCGA pipeline, GEO validation modules, fast-death sensitivity analyses "
    "(Supplementary Table S10), and GDC 2025 cross-validation is available at "
    "[INSERT PUBLIC REPOSITORY URL, e.g. https://github.com/YOUR_USERNAME/tcga-esca-prognostic-transfer] "
    "and archived at [INSERT ZENODO DOI ON FIRST SUBMISSION]. Primary reproduction commands: "
    "python3 -m esca_escc_multiomic_pipeline.run_all and "
    "python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity."
)

CODE_AVAILABILITY = (
    "Custom Python 3 scripts (pandas, numpy, scipy, lifelines, matplotlib, seaborn) are "
    "provided in the public repository named above. No restrictions on access."
)

REMARK_METHODS = (
    " Prognostic model reporting follows REMARK guidance (Altman et al., J Clin Oncol 2010); "
    "compliance items are listed in Supplementary Table S12."
)

MUTATION_RESULTS_PREFIX = (
    "To provide a mechanistic substrate for our lineage-resolved findings, we integrated an "
    "exploratory somatic layer investigating the interplay between mutations and histological "
    "backgrounds. Recognizing that mutation-by-histology interaction tests are often "
    "constrained by event counts in moderate sample sizes, we treated these analyses as a "
    "prioritization framework for identifying putative effect modifiers rather than definitive "
    "evidence of clinical modification. This integrated approach is summarized across the "
    "expression-module shifts in Fig. 4A, the prioritized associations in Supplementary Table S11, "
    "and the survival interaction landscape visualized in Fig. 4B."
)

S10_CAPTION = (
    "Supplementary Table S10. Fast-death death-tail sensitivity (10%, 20%, 30%) on the OS60 "
    "cohort. Fast-death = shortest-survival tail among OS60 decedents (minimum n=10). "
    "Methylation contrasts: fast-death vs all other samples within histology (Welch t-test; "
    "BH FDR). Patient membership Jaccard and top-200 probe Jaccard (lowest q_diff) vs the "
    "20% primary definition. Source: result/fast_death_sensitivity_summary.tsv; reproduced by "
    "python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity."
)

S11_CAPTION = (
    "Supplementary Table S11. OS60 genome-wide mutation×histology interaction screen "
    "(nominal Wald p-values; hypothesis-generating). Source: "
    "result/mutation_x_histology_interactions_OS.tsv."
)

S12_CAPTION = (
    "Supplementary Table S12. REMARK checklist for prognostic model reporting (Y = reported; "
    "P = partial; N = not applicable). Reference: Altman DG et al., J Clin Oncol 2010."
)

S10_TABLE_HEADERS = [
    "Histology",
    "Death-tail (%)",
    "n cohort",
    "n OS60 deaths",
    "n fast-death",
    "Max OS fast (mo)",
    "n probes q<0.05",
    "Patient Jaccard vs 20%",
    "Top-200 probe Jaccard vs 20%",
]

S10_TABLE_ROWS = [
    ["EAC/ESCA", "10", "89", "44", "10", "5.91", "158", "1.00", "1.00"],
    ["EAC/ESCA", "20", "89", "44", "10", "5.91", "158", "—", "—"],
    ["EAC/ESCA", "30", "89", "44", "14", "7.13", "62", "0.714", "0.311"],
    ["ESCC", "10", "96", "32", "10", "5.16", "68", "1.00", "1.00"],
    ["ESCC", "20", "96", "32", "10", "5.16", "68", "—", "—"],
    ["ESCC", "30", "96", "32", "10", "5.16", "68", "1.00", "1.00"],
]

S10_KEY_LOCUS_HEADERS = ["Histology", "Gene", "Rank @10%", "Rank @20%", "Rank @30%", "Stable?"]
S10_KEY_LOCUS_ROWS = [
    ["ESCC", "SHISA3", "1", "1", "1", "Yes"],
    ["ESCC", "C7orf13", "2", "2", "2", "Yes"],
    ["ESCC", "PRKG1", "3", "3", "3", "Yes"],
    ["ESCC", "SEMA3E", "4", "4", "4", "Yes"],
    ["EAC/ESCA", "SIGLEC14", "1", "1", "176", "No"],
    ["EAC/ESCA", "EPHB1", "2", "2", "81", "Partial"],
    ["EAC/ESCA", "ZNF214", "3", "3", "13533", "No"],
    ["EAC/ESCA", "RAB11FIP3", "5", "5", "11", "Partial"],
]

REMARK_HEADERS = ["Item", "Topic", "Status", "Where reported"]
REMARK_ROWS = [
    ["1", "Prognostic (not predictive) intent", "Y", "Introduction; Abstract"],
    ["2", "Study design", "Y", "Methods — TCGA secondary analysis"],
    ["3", "Setting and data source", "Y", "Methods — cBioPortal; GEO modules"],
    ["4", "Eligibility criteria", "Y", "Methods — primary tumours; EAC vs ESCC"],
    ["5", "Treatments", "P", "Introduction (CROSS); TCGA metadata"],
    ["6", "Outcome definition", "Y", "Methods — OS60 landmark"],
    ["7", "Predictor definition", "Y", "Methods — RNA + HM450 paired features"],
    ["8", "Sample size / events", "Y", "Results; Table 1 / S1"],
    ["9", "Patient flow diagram", "N", "Single public cohort"],
    ["10", "Patient characteristics", "P", "Not central to transfer conclusion"],
    ["11", "Univariable analysis", "P", "Table 6 / S6 (methylation–OS)"],
    ["12", "Multivariable model building", "Y", "Methods — elastic-net Cox"],
    ["13", "Model specification", "Y", "Supplementary Cox coefficient tables"],
    ["14", "Missing data", "Y", "Methods — median imputation"],
    ["15", "Events per variable", "P", "Penalized Cox; exploratory framing"],
    ["16", "Performance measures", "Y", "C-index; Table 1 / Fig. 1"],
    ["17", "Internal validation", "P", "GDC 2025 5-fold CV (Results)"],
    ["18", "External validation", "P", "GEO feasibility only; Limitations"],
    ["19", "Model comparison", "Y", "Table 1 — within / cross / pooled"],
    ["20", "Model presentation", "P", "Coefficients in TSV; no nomogram"],
    ["21", "Interpretation", "Y", "Discussion — histology-specific structure"],
    ["22", "Limitations", "Y", "Limitations section"],
    ["23", "Generalizability", "Y", "Discussion / Limitations"],
    ["24", "Conflict of interest", "Y", "Competing interests"],
    ["25", "Funding", "P", "Acknowledgements placeholder"],
]


def red_run(paragraph: Paragraph, text: str) -> None:
    r = paragraph.add_run(text)
    r.font.highlight_color = WD_COLOR_INDEX.RED


def red_replace(paragraph: Paragraph, text: str) -> None:
    paragraph.clear()
    red_run(paragraph, text)


def red_append(paragraph: Paragraph, text: str) -> None:
    red_run(paragraph, text)


def insert_red_after(paragraph: Paragraph, text: str, style: str | None = None) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style:
        new_para.style = style
    red_run(new_para, text)
    return new_para


def insert_table_after(
    paragraph: Paragraph,
    doc: Document,
    headers: list[str],
    rows: list[list[str]],
):
    n_rows = 1 + len(rows)
    n_cols = len(headers)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for j, h in enumerate(headers):
        hdr_cells[j].text = h
        for run in hdr_cells[j].paragraphs[0].runs:
            run.font.highlight_color = WD_COLOR_INDEX.RED
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = val
            for run in cell.paragraphs[0].runs:
                run.font.highlight_color = WD_COLOR_INDEX.RED
    paragraph._p.addnext(table._tbl)
    return table._tbl


def insert_red_after_element(
    element,
    parent,
    text: str,
    style: str | None = None,
) -> Paragraph:
    new_p = OxmlElement("w:p")
    element.addnext(new_p)
    new_para = Paragraph(new_p, parent)
    if style:
        new_para.style = style
    red_run(new_para, text)
    return new_para


def find_paragraph(doc: Document, prefix: str) -> Paragraph:
    for p in doc.paragraphs:
        if (p.text or "").startswith(prefix):
            return p
    raise RuntimeError(f"Not found: {prefix[:70]!r}")


def find_paragraph_contains(doc: Document, needle: str) -> Paragraph:
    for p in doc.paragraphs:
        if needle in (p.text or ""):
            return p
    raise RuntimeError(f"Not found containing: {needle[:70]!r}")


def main() -> None:
    doc = Document(DOC_PATH)

    # --- Results: fast-death sensitivity (fix Jaccard) ---
    red_replace(
        find_paragraph(doc, "Sensitivity of the fast-death definition."),
        RESULTS_SENSITIVITY,
    )

    # --- Methods: fast-death (fix Jaccard + S10 ref) ---
    red_replace(
        find_paragraph(doc, 'Rather than optimizing for broad prognostic signatures, we focused on "fast-death"'),
        METHODS_FAST_DEATH,
    )

    # --- Methods: REMARK ---
    p_cox = find_paragraph(
        doc,
        "A 60-month clinical landmark (OS60) was established for all survival analyses",
    )
    if "Supplementary Table S12" not in (p_cox.text or ""):
        red_append(p_cox, REMARK_METHODS)

    # --- Results: mutation S10 -> S11 ---
    red_replace(
        find_paragraph(doc, "To provide a mechanistic substrate for our lineage-resolved findings"),
        MUTATION_RESULTS_PREFIX,
    )

    # --- Data availability ---
    red_replace(
        find_paragraph_contains(doc, "TCGA esophageal carcinoma (ESCA) Firehose Legacy data were accessed"),
        DATA_AVAILABILITY,
    )

    # --- Code availability (insert after Data availability body if missing) ---
    if not any((p.text or "").startswith("Code availability") for p in doc.paragraphs):
        p_data_heading = find_paragraph(doc, "Data availability")
        p_data_body = find_paragraph_contains(
            doc, "TCGA esophageal carcinoma (ESCA) Firehose Legacy data (clinical annotations"
        )
        h = insert_red_after(p_data_body, "Code availability", style="Heading 1")
        insert_red_after(h, CODE_AVAILABILITY)

    # --- Supplementary materials blurb ---
    p_supp = find_paragraph_contains(doc, "Raw and processed outputs are stored alongside")
    if "S10–S12" not in (p_supp.text or ""):
        red_append(
            p_supp,
            " Supplementary Tables S10 (fast-death sensitivity), S11 (mutation×histology screen), "
            "and S12 (REMARK checklist) are described below.",
        )

    # --- Supplementary Table S10 / S11 / S12 ---
    p_s10 = find_paragraph(doc, "Supplementary Table S10.")
    red_replace(p_s10, S10_CAPTION)

    p_panel_a = insert_red_after(p_s10, "Panel A. Patient classification and probe-set stability.")
    tbl_a = insert_table_after(p_panel_a, doc, S10_TABLE_HEADERS, S10_TABLE_ROWS)

    p_panel_b = insert_red_after_element(
        tbl_a, p_s10._parent, "Panel B. Leading loci — genome-wide q_diff rank stability."
    )
    tbl_b = insert_table_after(p_panel_b, doc, S10_KEY_LOCUS_HEADERS, S10_KEY_LOCUS_ROWS)

    p_s11 = insert_red_after_element(tbl_b, p_s10._parent, S11_CAPTION)
    p_s12 = insert_red_after(p_s11, S12_CAPTION)
    insert_table_after(p_s12, doc, REMARK_HEADERS, REMARK_ROWS)

    doc.save(DOC_PATH)
    print(f"Saved submission-prep updates (red highlight) to:\n  {DOC_PATH}")


if __name__ == "__main__":
    main()
