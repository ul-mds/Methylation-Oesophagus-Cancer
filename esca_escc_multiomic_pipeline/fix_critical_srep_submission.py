"""
Fix Cancer Medicine critical submission blockers in Manuscript First.docx.
All new or replaced text is highlighted yellow (WD_COLOR_INDEX.YELLOW).

Fixes:
  1. Remove internal 'Additional information (display items)' editorial block
  2. Replace title-page / end-matter placeholders (affiliations, CRediT, funding, code URLs)
  3. Repair corrupted Figure 3 legend math (inline + Figure legends section)
  4. Add ethics statement to Methods
  5. Remove embedded figure images to bring manuscript under 3 MB (figures uploaded separately)
  6. Build Supplementary/Supplementary_Information.pdf
"""
from __future__ import annotations

from esca_escc_multiomic_pipeline.study_paths import study_root

import argparse
import io
import re
import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from fpdf import FPDF

STUDY_ROOT = study_root()
DEFAULT_DOCX = STUDY_ROOT / "Manuscript First.docx"
SUPP_DIR = STUDY_ROOT / "Supplementary"
SUPP_PDF = SUPP_DIR / "Supplementary_Information.pdf"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_DRAWING = f"{{{W_NS}}}drawing"
W_PICT = f"{{{W_NS}}}pict"
W_OBJECT = f"{{{W_NS}}}object"

FIG3_LEGEND = (
    "Figure 3. Fast-death differential methylation by histology (HM450). Genome-wide volcano plots of "
    "HM450 probes contrasting fast-death versus other deaths within each histology (left: ESCA/EAC; "
    "right: ESCC) on the OS60 landmark. The x-axis shows Δβ (mean β in fast-death minus mean β in "
    "other deaths) and the y-axis shows −log10(q_diff) from Benjamini–Hochberg FDR correction. "
    "(B) Ranked summaries of the top hypomethylated (Δβ < 0) and hypermethylated (Δβ > 0) loci for "
    "ESCA/EAC (left) and ESCC (right), with bar length indicating Δβ. Complete ranked probe lists "
    "are provided in Supplementary Tables S2–S5 online."
)
FIG3_INLINE = FIG3_LEGEND.replace("Figure 3.", "Fig. 3.", 1)

AFFIL_3 = (
    "3\tDepartment and Institution, City, Country (Rene Thieme — replace with confirmed affiliation "
    "before submission)"
)
AFFIL_4 = (
    "4\tDepartment and Institution, City, Country (Sanne Hoefnagel — replace with confirmed "
    "affiliation before submission)"
)

AUTHOR_CONTRIB = (
    "Senol Dogan: conceptualization, methodology, software, validation, formal analysis, investigation, "
    "data curation, writing—original draft, visualization, project administration. "
    "Toralf Kirsten: supervision, resources, writing—review and editing. "
    "Rene Thieme: writing—review and editing. "
    "Sanne Hoefnagel: writing—review and editing. "
    "All authors approved the final version."
)

ACKNOWLEDGEMENTS = (
    "The authors thank the TCGA Research Network and the cBioPortal team for curated public data. "
    "Funding: [insert grant agency and award number(s) before submission]. "
    "Main-article figures were generated using Python (matplotlib 3.8, https://matplotlib.org) and "
    "seaborn (0.13, https://seaborn.pydata.org) with source code in the public repository described "
    "under Code availability."
)

DATA_AVAILABILITY = (
    "TCGA esophageal carcinoma (ESCA) Firehose Legacy data (clinical annotations, HM450 methylation "
    "beta values, RNA expression z-scores, and mutation calls) were accessed through cBioPortal "
    "(https://www.cbioportal.org/study/summary?id=esca_tcga). External GEO series used for modular "
    "feasibility analyses were GSE161533 (GPL570; paired tumour–normal expression), GSE53624 "
    "(GPL18109; lncRNA microarray, training), and GSE53622 (GPL18109; lncRNA microarray, hold-out "
    "scoring). Processed tabular outputs supporting all figures and tables (including Supplementary "
    "Tables S1–S12) are in the project result/ directory; main-text figure images are in Plots/; "
    "supplementary tables and figures are in Supplementary/. Analysis code to reproduce the primary "
    "TCGA pipeline, GEO validation modules, fast-death sensitivity analyses (Supplementary Table S10), "
    "and GDC 2025 cross-validation will be deposited in a public GitHub repository and archived on "
    "Zenodo with a citable DOI before acceptance (insert final repository URL and "
    "https://doi.org/10.5281/zenodo.XXXXXXX before submission). Primary reproduction commands: "
    "python3 -m esca_escc_multiomic_pipeline.run_all and "
    "python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity."
)

ETHICS_METHODS = (
    "Ethics approval and consent to participate: Not applicable. This study re-analysed de-identified, "
    "publicly available TCGA and GEO data with no access to identifiable participant information."
)


def yellow_run(paragraph: Paragraph, text: str) -> None:
    run = paragraph.add_run(text)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def replace_paragraph_yellow(paragraph: Paragraph, text: str) -> None:
    paragraph.clear()
    yellow_run(paragraph, text)


def insert_paragraph_after(paragraph: Paragraph, text: str) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    yellow_run(new_para, text)
    return new_para


def find_paragraph(doc: Document, predicate) -> Paragraph:
    for p in doc.paragraphs:
        if predicate(p.text or ""):
            return p
    raise RuntimeError("Paragraph not found")


def remove_paragraph(paragraph: Paragraph) -> None:
    el = paragraph._element
    el.getparent().remove(el)


def remove_all_embedded_images(doc: Document) -> int:
    removed = 0
    for para in doc.paragraphs:
        for run in para.runs:
            for child in list(run._element):
                tag = child.tag.split("}")[-1]
                if tag in ("drawing", "pict", "object", "AlternateContent"):
                    run._element.remove(child)
                    removed += 1
    for el in doc.element.body.iter():
        if el.tag in (W_DRAWING, W_PICT, W_OBJECT):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
                removed += 1
    return removed


def fix_figure3_text(text: str) -> str | None:
    if "Fast-death differential methylation by histology (HM450)" not in text:
        return None
    if text.strip().startswith("Fig. 3."):
        return FIG3_INLINE
    if text.strip().startswith("Figure 3."):
        return FIG3_LEGEND
    return None


def remove_editorial_display_items_block(doc: Document) -> None:
    to_remove: list[Paragraph] = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == "Additional information (display items)":
            to_remove.append(p)
        elif t.startswith("This manuscript currently includes five figures and three tables"):
            to_remove.append(p)
        elif t.startswith("Cancer Medicine author guidelines note: at most eight combined figures"):
            to_remove.append(p)
    for p in to_remove:
        remove_paragraph(p)


def build_supplementary_pdf(doc: Document) -> None:
    title = doc.paragraphs[0].text.strip()
    authors = ""
    for p in doc.paragraphs[:15]:
        if "Dogan" in (p.text or "") and "*" in (p.text or ""):
            authors = p.text.strip()
            break

    supp_start = None
    for i, p in enumerate(doc.paragraphs):
        if (p.text or "").strip() == "Supplementary Information":
            supp_start = i
            break
    if supp_start is None:
        raise RuntimeError("Supplementary Information section not found")

    lines: list[str] = []
    for p in doc.paragraphs[supp_start + 1 :]:
        t = (p.text or "").strip()
        if t:
            lines.append(t)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 12)
    pdf.multi_cell(0, 7, title)
    pdf.ln(2)
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 5, authors)
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 6, "Supplementary Information", ln=True)
    pdf.ln(2)
    pdf.set_font("Helvetica", "", 9)

    for line in lines:
        if line == "Supplementary Tables":
            pdf.ln(2)
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 5, line, ln=True)
            pdf.set_font("Helvetica", "", 9)
            continue
        if line == "Supplementary Figures":
            pdf.ln(2)
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 5, line, ln=True)
            pdf.set_font("Helvetica", "", 9)
            continue
        safe = line.encode("latin-1", errors="replace").decode("latin-1")
        pdf.multi_cell(0, 4.5, safe)
        pdf.ln(1)

    fig_specs = [
        (SUPP_DIR / "Figures" / "S1_mediation_delta_volcano.png", "Supplementary Fig. S1"),
        (SUPP_DIR / "Figures" / "S2_top_histology_specific_mediation.png", "Supplementary Fig. S2"),
    ]
    for fig_path, label in fig_specs:
        if fig_path.exists():
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 6, label, ln=True)
            pdf.ln(2)
            pdf.image(str(fig_path), w=170)

    pdf.ln(4)
    pdf.set_font("Helvetica", "I", 8)
    pdf.multi_cell(
        0,
        4,
        "Full supplementary tables S1-S12 are provided in Supplementary_Tables.xlsx in the "
        "Supplementary/ folder.",
    )

    SUPP_DIR.mkdir(parents=True, exist_ok=True)
    pdf.output(str(SUPP_PDF))


def strip_docx_media(docx_path: Path) -> None:
    """Remove word/media/* and image relationships after inline drawings were deleted."""
    buf = io.BytesIO()
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            if name.startswith("word/media/"):
                continue
            data = zin.read(name)
            if name == "word/_rels/document.xml.rels":
                text = data.decode("utf-8")
                text = re.sub(
                    r'<Relationship[^>]+Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"[^>]*/>\s*',
                    "",
                    text,
                )
                data = text.encode("utf-8")
            elif name == "[Content_Types].xml":
                text = data.decode("utf-8")
                text = re.sub(r'<Default Extension="(png|jpeg|jpg|gif|emf|wmf)"[^>]*/>\s*', "", text)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    docx_path.write_bytes(buf.getvalue())


def apply_fixes(docx_path: Path) -> None:
    backup = docx_path.with_suffix(docx_path.suffix + ".pre_critical_fix.bak")
    shutil.copy2(docx_path, backup)

    doc = Document(str(docx_path))

    # --- Title page affiliations ---
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t.startswith("3") and "AFFILIATION" in t.upper():
            replace_paragraph_yellow(p, AFFIL_3)
        elif t.startswith("4") and "AFFILIATION" in t.upper():
            replace_paragraph_yellow(p, AFFIL_4)

    # --- End matter placeholders ---
    replace_paragraph_yellow(
        find_paragraph(doc, lambda t: t.startswith("The authors thank the TCGA")),
        ACKNOWLEDGEMENTS,
    )
    replace_paragraph_yellow(
        find_paragraph(doc, lambda t: "[assign CRediT roles]" in t),
        AUTHOR_CONTRIB,
    )
    replace_paragraph_yellow(
        find_paragraph(doc, lambda t: t.startswith("TCGA esophageal carcinoma") and "cbioportal" in t.lower()),
        DATA_AVAILABILITY,
    )

    # --- Ethics in Methods ---
    methods_p = find_paragraph(doc, lambda t: t.strip() == "Methods")
    meth_start = next(i for i, p in enumerate(doc.paragraphs) if (p.text or "").strip() == "Methods")
    refs_i = next(
        (i for i, p in enumerate(doc.paragraphs) if (p.text or "").strip() == "References"),
        len(doc.paragraphs),
    )
    methods_text = "\n".join((doc.paragraphs[i].text or "") for i in range(meth_start, refs_i))
    if "Ethics approval and consent to participate" not in methods_text:
        insert_paragraph_after(methods_p, ETHICS_METHODS)

    # --- Figure 3 legends ---
    for p in doc.paragraphs:
        fixed = fix_figure3_text(p.text or "")
        if fixed:
            replace_paragraph_yellow(p, fixed)

    # --- Remove editorial block ---
    remove_editorial_display_items_block(doc)

    # --- Remove embedded images (upload separately) ---
    n_removed = remove_all_embedded_images(doc)

    doc.save(str(docx_path))
    strip_docx_media(docx_path)
    size_mb = docx_path.stat().st_size / 1024 / 1024

    doc2 = Document(str(docx_path))
    build_supplementary_pdf(doc2)

    print(f"Backup: {backup}")
    print(f"Updated: {docx_path}")
    print(f"Removed {n_removed} embedded image element(s); manuscript size: {size_mb:.2f} MB")
    print(f"Supplementary PDF: {SUPP_PDF} ({SUPP_PDF.stat().st_size / 1024:.0f} KB)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Fix SR critical submission issues (yellow highlight).")
    parser.add_argument("--docx", type=Path, default=DEFAULT_DOCX)
    args = parser.parse_args()
    apply_fixes(args.docx)


if __name__ == "__main__":
    main()
