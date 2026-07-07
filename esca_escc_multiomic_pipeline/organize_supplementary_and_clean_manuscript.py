"""
Organize all supplementary figures/tables under Supplementary/ and clean Manuscript First.docx:

- Supplementary/Figures/ — S1–S2 PNGs only (mediation panels; NOT main-text figures)
- Supplementary/Tables/ — source TSV copies
- Supplementary/Supplementary_Figures.docx — S1–S2 with images (no main-figure duplicates)
- Supplementary/Supplementary_Tables.xlsx — S1–S12 sheets from result/ and GEO outputs

Manuscript (Scientific Reports):
- Main text keeps inline Figures 1–5 and Tables 2–4 only
- Supplementary Information section: titles/legends ONLY (no embedded table grids or figure images)
- Remove duplicate Fig. 3 stray image (combined volcano+bars panel)
"""
from __future__ import annotations

import argparse
import hashlib
import re
import shutil
from pathlib import Path

import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from lxml import etree

STUDY_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DOCX = STUDY_ROOT / "Manuscript First.docx"
SUPP_DIR = STUDY_ROOT / "Supplementary"
SUPP_FIG_DIR = SUPP_DIR / "Figures"
SUPP_TAB_DIR = SUPP_DIR / "Tables"
PLOTS = STUDY_ROOT / "Plots"
RESULT = STUDY_ROOT / "result"
GEO = STUDY_ROOT / "geo_gse536xx_validation"

W_T = qn("w:t")
W_P = qn("w:p")
W_TBL = qn("w:tbl")

MAIN_FIGURE_PNGS = {
    "00_analysis_pipeline_flowchart.png",
    "01_cross_apply_cindex.png",
    "02_fastdeath_methylation_volcano_panel.png",
    "02_fastdeath_combined_A_volcano_B_bars.png",
    "03_fastdeath_top_hypo_hyper_panel.png",
    "04_KM_methylation_60mo_panel.png",
    "05_mutation_module_heatmap.png",
    "05_somatic_combined_A_heatmap_B_interaction.png",
    "06_interaction_screen.png",
}

SUPP_FIGURE_SPECS: list[tuple[str, str, str]] = [
    (
        "S1_mediation_delta_volcano.png",
        "07_mediation_delta_volcano.png",
        "Supplementary Fig. S1. Mediation-style delta landscape (volcano): ESCC versus EAC/ESCA "
        "(supplementary to main mediation analyses).",
    ),
    (
        "S2_top_histology_specific_mediation.png",
        "08_top_histology_specific_mediation.png",
        "Supplementary Fig. S2. Top histology-specific mediation contrasts (supplementary summary figure).",
    ),
]

TABLE_SOURCES: list[tuple[str, str, Path | None]] = [
    ("S1", "Concordance index (C-index) for histology-specific and pooled models on overall survival.", RESULT / "model_performance_cindex.tsv"),
    ("S2", "EAC/ESCA: top hypomethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).", RESULT / "methylation_fast_death_EAC_ESCA.tsv"),
    ("S3", "ESCC: top hypomethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).", RESULT / "methylation_fast_death_ESCC.tsv"),
    ("S4", "EAC/ESCA: top hypermethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).", RESULT / "methylation_fast_death_EAC_ESCA.tsv"),
    ("S5", "ESCC: top hypermethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).", RESULT / "methylation_fast_death_ESCC.tsv"),
    ("S6", "EAC/ESCA: univariate Cox associations between continuous probe methylation and overall survival (selected top probes by q).", RESULT / "methylation_fast_death_EAC_ESCA.tsv"),
    ("S7", "Exploratory mechanistic follow-up: strongest associations between somatic mutation and expression or module scores, stratified by histology (ranked by nominal p-value).", RESULT / "interaction_mechanism_mut_to_expr_modules.tsv"),
    ("S8", "GSE161533 (GEO GPL570): TCGA pre-specified ESCC gene union overlapping tumour–normal expression summaries.", GEO / "gse161533_tcga_gene_overlap_detail.tsv"),
    ("S9", "GSE53624 / GSE53622: landmark (60-month) penalized Cox on lncRNA features.", GEO / "validation_landmark60_report.tsv"),
    ("S10_summary", "Fast-death death-tail sensitivity — cohort summary (10%, 20%, 30%).", RESULT / "fast_death_sensitivity_summary.tsv"),
    ("S10_key_loci", "Fast-death death-tail sensitivity — leading locus rank stability.", RESULT / "fast_death_sensitivity_key_loci.tsv"),
    ("S11", "OS60 genome-wide mutation×histology interaction screen (nominal Wald p-values; hypothesis-generating).", RESULT / "mutation_x_histology_interactions_OS.tsv"),
]

REMARK_ROWS = [
    ["1", "Prognostic (not predictive) intent", "Y", "Introduction; Abstract"],
    ["2", "Study design", "Y", "Methods — TCGA secondary analysis"],
    ["3", "Setting and data source", "Y", "Methods — cBioPortal; GEO modules"],
    ["4", "Eligibility criteria", "Y", "Methods — primary tumours; EAC vs ESCC"],
    ["5", "Treatments", "P", "Introduction (CROSS); TCGA metadata"],
    ["6", "Outcome definition", "Y", "Methods — OS60 landmark"],
    ["7", "Predictor definition", "Y", "Methods — RNA + HM450 paired features"],
    ["8", "Sample size / events", "Y", "Results; Supplementary Table S1"],
    ["9", "Patient flow diagram", "N", "Single public cohort"],
    ["10", "Patient characteristics", "P", "Not central to transfer conclusion"],
    ["11", "Univariable analysis", "P", "Table 4 / Supplementary Table S6"],
    ["12", "Multivariable model building", "Y", "Methods — elastic-net Cox"],
    ["13", "Model specification", "Y", "Supplementary Cox coefficient tables"],
    ["14", "Missing data", "Y", "Methods — median imputation"],
    ["15", "Events per variable", "P", "Penalized Cox; exploratory framing"],
    ["16", "Performance measures", "Y", "C-index; Supplementary Table S1 / Fig. 2"],
    ["17", "Internal validation", "P", "GDC 2025 5-fold CV (Results)"],
    ["18", "External validation", "P", "GEO feasibility only; Discussion"],
    ["19", "Model comparison", "Y", "Supplementary Table S1 — within / cross / pooled"],
    ["20", "Model presentation", "P", "Coefficients in TSV; no nomogram"],
    ["21", "Interpretation", "Y", "Discussion — histology-specific structure"],
    ["22", "Limitations", "Y", "Discussion"],
    ["23", "Generalizability", "Y", "Discussion"],
    ["24", "Conflict of interest", "Y", "Competing interests"],
    ["25", "Funding", "P", "Acknowledgements placeholder"],
]

SUPP_INTRO = (
    "Supplementary Tables S1–S12 and Supplementary Figs. S1–S2 are provided as separate files in the "
    "Supplementary/ folder (Supplementary_Tables.xlsx and Supplementary_Figures.docx; PNGs in "
    "Supplementary/Figures/). Per Scientific Reports guidelines, full tabular data and supplementary "
    "figure images are not embedded in the main manuscript; brief titles and legends are listed below."
)

S10_CAPTION = (
    "Supplementary Table S10. Fast-death death-tail sensitivity (10%, 20%, 30%) on the OS60 cohort. "
    "Fast-death = shortest-survival tail among OS60 decedents (minimum n=10). Methylation contrasts: "
    "fast-death vs all other samples within histology (Welch t-test; BH FDR). Patient membership Jaccard "
    "and top-200 probe Jaccard (lowest q_diff) vs the 20% primary definition. Full data: "
    "Supplementary/Tables/S10_fast_death_sensitivity_summary.tsv and S10_fast_death_sensitivity_key_loci.tsv."
)


def _ptext(el) -> str:
    return "".join(t.text or "" for t in el.iter(W_T)).strip()


def _p_has_image(p: Paragraph) -> bool:
    return bool(p._p.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))


def _plot_hash(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _embedded_png_name(p: Paragraph, plot_hashes: dict[str, str]) -> str | None:
    for blip in p._p.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
        r_id = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        if not r_id:
            continue
        blob = p.part.related_parts[r_id].blob
        return plot_hashes.get(hashlib.md5(blob).hexdigest())
    return None


def setup_supplementary_folder() -> None:
    SUPP_FIG_DIR.mkdir(parents=True, exist_ok=True)
    SUPP_TAB_DIR.mkdir(parents=True, exist_ok=True)

    for dest_name, src_name, _ in SUPP_FIGURE_SPECS:
        src = PLOTS / src_name
        dest = SUPP_FIG_DIR / dest_name
        if not src.is_file():
            raise FileNotFoundError(src)
        shutil.copy2(src, dest)

    for sheet, _title, src in TABLE_SOURCES:
        if src is None or not src.is_file():
            print(f"Warning: missing table source for {sheet}: {src}")
            continue
        dest = SUPP_TAB_DIR / f"{sheet}_{src.name}"
        shutil.copy2(src, dest)

    xlsx_path = SUPP_DIR / "Supplementary_Tables.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for sheet, title, src in TABLE_SOURCES:
            if src is None or not src.is_file():
                pd.DataFrame({"note": [f"Source file pending: {title}"]}).to_excel(
                    writer, sheet_name=sheet[:31], index=False
                )
                continue
            df = pd.read_csv(src, sep="\t")
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
        pd.DataFrame(REMARK_ROWS, columns=["Item", "Topic", "Status", "Where reported"]).to_excel(
            writer, sheet_name="S12", index=False
        )
    print(f"Wrote {xlsx_path}")

    fig_doc_path = SUPP_DIR / "Supplementary_Figures.docx"
    if fig_doc_path.is_file():
        shutil.copy2(fig_doc_path, fig_doc_path.with_suffix(".docx.old_bak"))
    fig_doc = Document()
    fig_doc.add_heading("Supplementary Figures", level=1)
    fig_doc.add_paragraph(
        "Mediation-focused panels not reproduced in the main article (Figures 1–5). "
        "Main-text figure PNGs are excluded from this file."
    )
    for dest_name, _src_name, caption in SUPP_FIGURE_SPECS:
        png = SUPP_FIG_DIR / dest_name
        fig_doc.add_paragraph()
        fig_doc.add_picture(str(png), width=Inches(6.2))
        p = fig_doc.add_paragraph()
        r = p.add_run(caption)
        r.bold = True
    fig_doc.save(str(fig_doc_path))
    print(f"Wrote {fig_doc_path} ({len(SUPP_FIGURE_SPECS)} unique supplementary figures)")


def _remove_fig3_stray_image(doc: Document) -> int:
    plot_hashes = {_plot_hash(p): p.name for p in PLOTS.glob("*.png")}
    removed = 0
    for i, p in enumerate(doc.paragraphs):
        if not _p_has_image(p):
            continue
        name = _embedded_png_name(p, plot_hashes)
        if name != "02_fastdeath_combined_A_volcano_B_bars.png":
            continue
        nxt = (doc.paragraphs[i + 1].text or "").strip() if i + 1 < len(doc.paragraphs) else ""
        if nxt.startswith("Fig. 3."):
            continue
        p._p.getparent().remove(p._p)
        removed += 1
    return removed


def _remove_supplementary_embedded_content(doc: Document) -> tuple[int, int]:
    """Remove table grids and figure images from Supplementary Information section."""
    body = doc.element.body
    in_supp = False
    past_fig_heading = False
    tables_removed = 0
    images_removed = 0

    for el in list(body):
        if el.tag == W_P:
            t = _ptext(el)
            if t == "Supplementary Information":
                in_supp = True
                continue
            if in_supp and t == "Supplementary Figures":
                past_fig_heading = True
                continue
            if in_supp and past_fig_heading and _p_has_image(Paragraph(el, doc._body)):
                el.getparent().remove(el)
                images_removed += 1
        elif el.tag == W_TBL and in_supp:
            el.getparent().remove(el)
            tables_removed += 1

    return tables_removed, images_removed


def _set_paragraph_text(p: Paragraph, new_text: str) -> None:
    ts = list(p._p.iter(W_T))
    if not ts:
        p.add_run(new_text)
        return
    ts[0].text = new_text
    for t in ts[1:]:
        t.text = ""


def _clean_supplementary_legends(doc: Document) -> None:
    in_supp = False
    remove_els = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == "Supplementary Information":
            in_supp = True
            continue
        if not in_supp:
            continue
        if t.startswith("Raw and processed outputs are stored"):
            _set_paragraph_text(p, SUPP_INTRO)
        if t.startswith("Supplementary Fig.s"):
            _set_paragraph_text(p, p.text.replace("Supplementary Fig.s", "Supplementary Figs."))
        if t in ("Panel A. Patient classification and probe-set stability.", "Panel B. Leading loci — genome-wide q_diff rank stability."):
            remove_els.append(p._p)
        if t.startswith("Supplementary Table S10."):
            _set_paragraph_text(p, S10_CAPTION)
        if "Source: result/" in t and t.startswith("Supplementary Table"):
            _set_paragraph_text(p, re.sub(r"\s*Source: result/[^.]+\.", ".", t))
    for el in remove_els:
        el.getparent().remove(el)


def _update_data_availability(doc: Document) -> None:
    for p in doc.paragraphs:
        t = p.text or ""
        if not t.startswith("TCGA esophageal carcinoma") or "Supplementary/" in t:
            continue
        new = t.replace(
            "figure images are in Plots/.",
            "main-text figure images are in Plots/; supplementary tables and figures are in Supplementary/.",
        )
        if new != t:
            _set_paragraph_text(p, new)


def _audit(doc: Document) -> list[str]:
    issues: list[str] = []
    plot_hashes = {_plot_hash(p): p.name for p in PLOTS.glob("*.png")}
    in_supp = False

    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t == "Supplementary Information":
            in_supp = True
        if in_supp and _p_has_image(p):
            name = _embedded_png_name(p, plot_hashes)
            issues.append(f"Supplementary section still has embedded image para {i}: {name}")
        if _p_has_image(p):
            name = _embedded_png_name(p, plot_hashes)
            if in_supp and name in MAIN_FIGURE_PNGS:
                issues.append(f"Main figure duplicated in supplementary: {name}")

    for tbl in doc.tables:
        if len(doc.tables) > 3:
            pass
    n_tables = len(doc.tables)
    if n_tables != 3:
        issues.append(f"Expected 3 main-text tables, found {n_tables}")

    main_imgs = []
    for i, p in enumerate(doc.paragraphs):
        if _p_has_image(p):
            name = _embedded_png_name(p, plot_hashes)
            main_imgs.append((i, name))
    if any(n == "02_fastdeath_combined_A_volcano_B_bars.png" for _, n in main_imgs):
        issues.append("Stray combined Fig. 3 panel still embedded")

    supp_fig_doc = SUPP_DIR / "Supplementary_Figures.docx"
    if supp_fig_doc.is_file():
        sdoc = Document(str(supp_fig_doc))
        for p in sdoc.paragraphs:
            if not _p_has_image(p):
                continue
            name = _embedded_png_name(p, plot_hashes)
            if name in MAIN_FIGURE_PNGS:
                issues.append(f"Supplementary_Figures.docx contains main figure: {name}")

    return issues


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", type=Path, default=DEFAULT_DOCX)
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()
    path = args.docx.resolve()

    if not args.no_backup:
        bak = path.with_suffix(path.suffix + ".pre_supp_folder.bak")
        shutil.copy2(path, bak)
        print(f"Backup: {bak}")

    setup_supplementary_folder()

    doc = Document(str(path))
    n_stray = _remove_fig3_stray_image(doc)
    n_tbl, n_img = _remove_supplementary_embedded_content(doc)
    _clean_supplementary_legends(doc)
    _update_data_availability(doc)
    doc.save(str(path))

    issues = _audit(Document(str(path)))
    print(f"Removed stray Fig.3 image: {n_stray}; supplementary tables: {n_tbl}; supplementary images: {n_img}")
    if issues:
        print("Audit warnings:")
        for x in issues:
            print(f"  - {x}")
    else:
        print("Audit OK: supplementary content in Supplementary/ folder; manuscript legends only.")
    print(f"Updated: {path}")


if __name__ == "__main__":
    main()
