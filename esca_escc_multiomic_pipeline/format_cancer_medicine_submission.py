"""Reformat the BMC Cancer submission package to Cancer Medicine (Wiley) layout.

Format only: section order, journal labels, table abbreviation footnotes,
supporting-information file names, and cover-letter addressee.
Body sentences of the manuscript are not rewritten.
https://onlinelibrary.wiley.com/page/journal/20457634/homepage/forauthors.html
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

from esca_escc_multiomic_pipeline.study_paths import (
    pipeline_dir,
    study_root,
    submission_package_dir,
)

ROOT = study_root()
PKG = submission_package_dir()
PIPE = pipeline_dir()
sys.path.insert(0, str(PIPE))

import build_bmc_additional_files as bmc  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REVIEW_HIGHLIGHT = os.environ.get("CM_REVIEW_HIGHLIGHT") == "1"


def ptext(el) -> str:
    return "".join(t.text or "" for t in el.findall(f".//{{{W}}}t")).strip()


def find_p(doc: Document, exact: str) -> Paragraph:
    for p in doc.paragraphs:
        if p.text.strip() == exact:
            return p
    raise ValueError(f"Paragraph not found: {exact!r}")


def find_startswith(doc: Document, prefix: str) -> Paragraph:
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefix):
            return p
    raise ValueError(f"Paragraph starting with {prefix!r} not found")


def set_label(paragraph: Paragraph, text: str) -> None:
    if not paragraph.runs:
        paragraph.add_run(text)
        return
    paragraph.runs[0].text = text
    for run in paragraph.runs[1:]:
        run.text = ""


def replace_once(paragraph: Paragraph, old: str, new: str) -> None:
    full = paragraph.text
    if old not in full:
        raise ValueError(f"Text not found: {old!r}")
    set_label(paragraph, full.replace(old, new, 1))


def block(start: Paragraph, end: Paragraph) -> list:
    out = []
    el = start._element
    end_el = end._element
    while el is not None:
        out.append(el)
        if el is end_el:
            return out
        el = el.getnext()
    raise ValueError("End element not reached")


def move_before(elements: list, anchor) -> None:
    for el in elements:
        anchor.addprevious(el)


def insert_after(paragraph: Paragraph, text: str) -> Paragraph:
    new_el = paragraph._p.makeelement(qn("w:p"), {})
    paragraph._p.addnext(new_el)
    new_p = Paragraph(new_el, paragraph._parent)
    new_p.style = "Normal"
    run = new_p.add_run(text)
    run.font.size = Pt(10)
    run.italic = True
    return new_p


def bold_header(table) -> None:
    for cell in table.rows[0].cells:
        for para in cell.paragraphs:
            if not para.runs:
                para.add_run(cell.text)
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(10)


def rename_header(table, old: str, new: str) -> None:
    for cell in table.rows[0].cells:
        if cell.text.strip() == old:
            para = cell.paragraphs[0]
            set_label(para, new)
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(10)
            return
    raise ValueError(f"Header {old!r} not in {[c.text for c in table.rows[0].cells]}")


def reformat_manuscript() -> None:
    path = PKG / "02_Manuscript.docx"
    doc = Document(str(path))

    ack_h = find_p(doc, "Acknowledgements")
    ack_b = ack_h._element.getnext()
    fund_h = find_p(doc, "Funding")
    fund_b = fund_h._element.getnext()
    eth_h = find_p(doc, "Ethics approval and consent to participate")
    eth_b = eth_h._element.getnext()
    con_h = find_p(doc, "Consent for publication")
    con_b = con_h._element.getnext()
    data_h = find_p(doc, "Availability of data and materials")
    data_b1 = data_h._element.getnext()
    data_b2 = data_b1.getnext()
    coi_h = find_p(doc, "Competing interests")
    coi_b = coi_h._element.getnext()
    au_h = find_p(doc, "Authors' contributions")
    au_b = au_h._element.getnext()
    decl = find_p(doc, "Declarations")
    abstract = find_p(doc, "Abstract")

    title_page = [
        ack_h._element, ack_b,
        fund_h._element, fund_b,
        eth_h._element, eth_b,
        con_h._element, con_b,
        data_h._element, data_b1, data_b2,
        coi_h._element, coi_b,
        au_h._element, au_b,
    ]
    move_before(title_page, abstract._element)
    decl._element.getparent().remove(decl._element)

    set_label(ack_h, "Acknowledgements")
    set_label(fund_h, "Funding statement")
    set_label(eth_h, "Ethical approval statement")
    set_label(data_h, "Data availability statement")
    set_label(coi_h, "Conflict of interest statement")
    set_label(au_h, "Author contributions")
    for heading in (ack_h, fund_h, eth_h, con_h, data_h, coi_h, au_h):
        heading.style = "Heading 2"

    set_label(find_p(doc, "Background"), "Introduction")
    set_label(find_p(doc, "References"), "Literature cited")
    set_label(find_p(doc, "Additional files"), "Supporting Information")
    replace_once(find_startswith(doc, "Additional file 1."), "Additional file 1.", "Supporting Information file 1.")
    replace_once(find_startswith(doc, "Additional file 2."), "Additional file 2.", "Supporting Information file 2.")

    abbrev = find_startswith(doc, "EAC: esophageal adenocarcinoma")
    abbrev.style = "Normal"

    junk = find_p(doc, "Formularende")
    junk._element.getparent().remove(junk._element)

    for p in list(doc.paragraphs):
        if p.style.name == "Heading 1" and not p.text.strip():
            p._element.getparent().remove(p._element)

    tables_h = find_p(doc, "Tables")
    table3_legend = find_startswith(doc, "Univariate Cox models of continuous HM450")
    figleg_h = find_p(doc, "Figure legends")
    si_h = find_p(doc, "Supporting Information")
    fig5 = find_startswith(doc, "Figure 5.")
    si_end = figleg_h._element.getprevious()

    move_before(block(tables_h, table3_legend), si_h._element)
    move_before(block(figleg_h, fig5), si_h._element)
    # si_end was the node before Figure legends; after the move it still precedes SI.
    _ = si_end

    footnotes = [
        (
            "Fast-death is the shortest-survival tail among EAC/ESCA decedents",
            "†β, methylation beta value. ‡Δβ, mean β (fast-death) − mean β (other). §q, Benjamini–Hochberg false discovery rate.",
        ),
        (
            "Definitions of fast-death, Δβ and q follow Table 1",
            "†β, methylation beta value. ‡Δβ, mean β (fast-death) − mean β (other). §q, Benjamini–Hochberg false discovery rate. Symbols match Table 1.",
        ),
        (
            "Univariate Cox models of continuous HM450",
            "†Cox β, log-hazard per unit increase in methylation beta. ‡q, Benjamini–Hochberg false discovery rate. §n, sample size. ¶Events, number of OS60 events.",
        ),
    ]
    for prefix, note in footnotes:
        insert_after(find_startswith(doc, prefix), note)

    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith("Table ") and p.style.name == "Normal" and len(t) < 120:
            for run in p.runs:
                run.bold = True

    rename_header(doc.tables[0], "Mean β fast-death", "Mean β fast-death†")
    rename_header(doc.tables[0], "Mean β other", "Mean β other†")
    rename_header(doc.tables[0], "Δβ", "Δβ‡")
    rename_header(doc.tables[0], "q", "q§")
    rename_header(doc.tables[1], "Mean β fast-death", "Mean β fast-death†")
    rename_header(doc.tables[1], "Mean β other", "Mean β other†")
    rename_header(doc.tables[1], "Δβ", "Δβ‡")
    rename_header(doc.tables[1], "q", "q§")
    rename_header(doc.tables[2], "Cox β", "Cox β†")
    rename_header(doc.tables[2], "q", "q‡")
    rename_header(doc.tables[2], "n", "n§")
    rename_header(doc.tables[2], "Events", "Events¶")
    for table in doc.tables:
        bold_header(table)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15
    pf.space_after = Pt(6)

    doc.save(str(path))


def reformat_cover_letter() -> None:
    src = PKG / "01_Cover_letter_BMC_Cancer.docx"
    doc = Document(str(src))
    set_label(doc.paragraphs[2], "Cancer Medicine Editorial Office")
    set_label(doc.paragraphs[3], "cancermed@wiley.com\t18 September 2026")
    replace_once(doc.paragraphs[4], "Research article —", "Original Research Article —")
    set_label(doc.paragraphs[5], "Dear Cancer Medicine Editorial Office,")
    replace_once(
        doc.paragraphs[6],
        "Please consider our Research article for BMC Cancer.",
        "Please consider our Original Research Article for Cancer Medicine.",
    )
    replace_once(
        doc.paragraphs[8],
        "We have read the Research article criteria, including the statement that manuscripts consisting solely of public-database computation without accompanying validation are not considered. ",
        "We have read the Original Research Article criteria for Cancer Medicine. ",
    )
    replace_once(
        doc.paragraphs[9],
        "and is why we believe the article belongs in BMC Cancer’s genetics/genomics/epigenetics and computational oncology scope.",
        "and is why we believe the article belongs in Cancer Medicine, which publishes bioinformatics, epigenetics, and prognostic research.",
    )
    replace_once(
        doc.paragraphs[15],
        "any additional files required",
        "any supporting information required",
    )
    out = PKG / "01_Cover_letter_Cancer_Medicine.docx"
    doc.save(str(out))
    src.unlink()


def _wiley_new_page(pdf, title: str | None = None):
    fig = plt.figure(figsize=bmc.A4)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.text(
        bmc.MARGIN / bmc.A4[0],
        1 - 0.32 / bmc.A4[1],
        "Supporting Information  ·  Cancer Medicine",
        fontsize=8,
        color="#444444",
    )
    ax.text(
        1 - bmc.MARGIN / bmc.A4[0],
        1 - 0.32 / bmc.A4[1],
        "Dogan et al.",
        fontsize=8,
        color="#444444",
        ha="right",
    )
    ax.plot(
        [bmc.MARGIN / bmc.A4[0], 1 - bmc.MARGIN / bmc.A4[0]],
        [1 - 0.42 / bmc.A4[1], 1 - 0.42 / bmc.A4[1]],
        color="#1f4e79",
        lw=0.8,
    )
    if title:
        ax.text(
            bmc.MARGIN / bmc.A4[0],
            1 - 0.72 / bmc.A4[1],
            title,
            fontsize=12,
            fontweight="bold",
            va="top",
            color="#1f4e79",
        )
    return fig, ax


def _wiley_cover(pdf, page_counter: list[int]):
    fig, ax = bmc.new_page(pdf)
    page_counter[0] += 1
    ax.text(bmc.MARGIN / bmc.A4[0], 0.90, "Supporting Information", fontsize=18, fontweight="bold", color="#1f4e79")
    ax.text(
        bmc.MARGIN / bmc.A4[0],
        0.85,
        bmc.wrap(
            "Histology-dependent prognostic model transportability and candidate methylation "
            "patterns in esophageal carcinoma",
            88,
        ),
        fontsize=11,
        va="top",
        **({"bbox": {"facecolor": "#fff59d", "edgecolor": "none", "pad": 2}} if REVIEW_HIGHLIGHT else {}),
    )
    ax.text(
        bmc.MARGIN / bmc.A4[0],
        0.74,
        bmc.wrap(
            "Senol Dogan, Sanne Hoefnagel, Chuang Yang, Rene Thieme, Toralf Kirsten. "
            "Cancer Medicine Original Research Article — supporting tables and figures.",
            100,
        ),
        fontsize=9,
        va="top",
        **({"bbox": {"facecolor": "#fff59d", "edgecolor": "none", "pad": 2}} if REVIEW_HIGHLIGHT else {}),
    )
    ax.text(
        bmc.MARGIN / bmc.A4[0],
        0.66,
        bmc.wrap(
            "This file presents every supplementary table (S1–S12) and supplementary figure (S1–S2) "
            "in the order they are first cited in the manuscript. Direction-specific gene-level methylation screens "
            "(S2–S5) and the methylation–OS60 Cox screen (S6) are shown as top-ranked previews; "
            "complete tables are in Supporting Information "
            "file 2 (.xlsx). Legends define abbreviations used in each table.",
            108,
        ),
        fontsize=9,
        va="top",
    )
    ax.text(bmc.MARGIN / bmc.A4[0], 0.52, "Contents (first-citation order)", fontsize=12, fontweight="bold", color="#1f4e79")
    labels = {
        "S12": "Table S12 — REMARK checklist",
        "F1": "Figure S1 — Mediation-style delta volcano",
        "F2": "Figure S2 — Top histology-specific mediation contrasts",
        "S2": "Table S2 — EAC gene-level short-survival hypomethylation (top 20)",
        "S3": "Table S3 — ESCC gene-level short-survival hypomethylation (top 20)",
        "S4": "Table S4 — EAC gene-level short-survival hypermethylation (top 20)",
        "S5": "Table S5 — ESCC gene-level short-survival hypermethylation (top 20)",
        "S10": "Table S10 — Short-survival 10/20/30% sensitivity",
        "S11": "Table S11 — Mutation×histology interaction screen (top 20)",
        "S7": "Table S7 — Mutation–expression/module shifts (top 20)",
        "S8": "Table S8 — GSE161533 gene overlap",
        "S9": "Table S9 — GSE53624/GSE53622 lncRNA Cox",
        "S1": "Table S1 — Cross-apply C-index matrix",
        "S6": "Table S6 — EAC methylation–OS60 Cox (top 20)",
    }
    y = 0.48
    for i, k in enumerate(bmc.CITE_ORDER, 1):
        ax.text(bmc.MARGIN / bmc.A4[0], y, f"{i:02d}.  {labels[k]}", fontsize=9, family="DejaVu Sans", va="top")
        y -= 0.026
    ax.text(
        bmc.MARGIN / bmc.A4[0],
        0.08,
        "Numerical order S1–S12 then Figures S1–S2 is preserved in Supporting Information file 2 sheet order.",
        fontsize=8,
        color="#444444",
    )
    bmc.footer(ax, page_counter[0])
    pdf.savefig(fig)
    plt.close(fig)


def build_supporting_information() -> None:
    for key, text in list(bmc.LEGENDS.items()):
        bmc.LEGENDS[key] = (
            text.replace("Additional file 2", "Supporting Information file 2")
            .replace("full Additional file 2", "Supporting Information file 2")
        )
    bmc.new_page = _wiley_new_page
    bmc.cover_and_contents = _wiley_cover

    def add_table_pages(pdf, key, page_counter):
        df = bmc.load_preview(key)
        title = f"Supporting Information — Table {key}. {bmc.SHORT_TITLES[key]}"
        legend = bmc.LEGENDS[key]
        max_rows = 22 if key != "S10" else 16
        n_chunks = max(1, __import__("math").ceil(len(df) / max_rows))
        for i in range(n_chunks):
            chunk = df.iloc[i * max_rows : (i + 1) * max_rows]
            head = title if i == 0 else title + f" (continued {i + 1}/{n_chunks})"
            fig, ax = bmc.new_page(pdf, head)
            page_counter[0] += 1
            y = 0.88
            if i == 0:
                kwargs = {"fontsize": 8, "va": "top"}
                if REVIEW_HIGHLIGHT and key in {"S2", "S3", "S4", "S5", "S6", "S12"}:
                    kwargs["bbox"] = {"facecolor": "#fff59d", "edgecolor": "none", "pad": 1.5}
                ax.text(bmc.MARGIN / bmc.A4[0], y, bmc.wrap("Legend. " + legend, 108), **kwargs)
                y = 0.88 - 0.018 * (bmc.wrap("Legend. " + legend, 108).count("\n") + 2)
            if key in ("S2", "S3", "S4", "S5", "S6", "S7", "S11"):
                note = f"Displaying {len(chunk)} of the ranked rows; the complete table is in Supporting Information file 2."
            else:
                note = f"Rows {i * max_rows + 1}–{i * max_rows + len(chunk)} of {len(df)}."
            ax.text(bmc.MARGIN / bmc.A4[0], y, note, fontsize=7.5, color="#333333", va="top")
            y -= 0.03
            bmc.draw_table_block(ax, chunk, y)
            bmc.footer(ax, page_counter[0])
            pdf.savefig(fig)
            plt.close(fig)

    def add_s10_pages(pdf, page_counter):
        title = f"Supporting Information — Table S10. {bmc.SHORT_TITLES['S10']}"
        legend = bmc.LEGENDS["S10"]
        import pandas as pd
        parts = [
            ("Panel A. Cohort membership and Jaccard overlap.", pd.read_excel(bmc.XLSX_SRC, sheet_name="S10_summary")),
            ("Panel B. Key loci across 10/20/30% death-tail thresholds.", pd.read_excel(bmc.XLSX_SRC, sheet_name="S10_key_loci")),
        ]
        for pi, (subtitle, df) in enumerate(parts):
            max_rows = 18
            n_chunks = max(1, __import__("math").ceil(len(df) / max_rows))
            for i in range(n_chunks):
                chunk = df.iloc[i * max_rows : (i + 1) * max_rows]
                head = title if pi == 0 and i == 0 else title + " (continued)"
                fig, ax = bmc.new_page(pdf, head)
                page_counter[0] += 1
                y = 0.88
                if pi == 0 and i == 0:
                    ax.text(bmc.MARGIN / bmc.A4[0], y, bmc.wrap("Legend. " + legend, 108), fontsize=8, va="top")
                    y -= 0.018 * (bmc.wrap("Legend. " + legend, 108).count("\n") + 2)
                ax.text(bmc.MARGIN / bmc.A4[0], y, subtitle, fontsize=9, fontweight="bold", va="top", color="#1f4e79")
                y -= 0.035
                bmc.draw_table_block(ax, chunk, y)
                bmc.footer(ax, page_counter[0])
                pdf.savefig(fig)
                plt.close(fig)

    def add_figure_page(pdf, key, img_path, page_counter):
        title = f"Supporting Information — Figure {key.replace('F', 'S')}. {bmc.SHORT_TITLES[key]}"
        fig, ax = bmc.new_page(pdf, title)
        page_counter[0] += 1
        ax.text(bmc.MARGIN / bmc.A4[0], 0.88, bmc.wrap("Legend. " + bmc.LEGENDS[key], 108), fontsize=8, va="top")
        from PIL import Image as PILImage
        im = PILImage.open(img_path).convert("RGB")
        left = bmc.MARGIN / bmc.A4[0]
        right = 1 - bmc.MARGIN / bmc.A4[0]
        bottom = 0.08
        top = 0.78
        box_w = (right - left) * bmc.A4[0]
        box_h = (top - bottom) * bmc.A4[1]
        iw, ih = im.size
        aspect = iw / ih
        box_aspect = box_w / box_h
        if aspect > box_aspect:
            w = right - left
            h = w * bmc.A4[0] / aspect / bmc.A4[1]
            x0 = left
            y0 = bottom + ((top - bottom) - h) / 2
        else:
            h = top - bottom
            w = h * bmc.A4[1] * aspect / bmc.A4[0]
            x0 = left + ((right - left) - w) / 2
            y0 = bottom
        ax_im = fig.add_axes([x0, y0, w, h])
        ax_im.imshow(im)
        ax_im.axis("off")
        bmc.footer(ax, page_counter[0])
        pdf.savefig(fig, dpi=300)
        plt.close(fig)

    bmc.add_table_pages = add_table_pages
    bmc.add_s10_pages = add_s10_pages
    bmc.add_figure_page = add_figure_page

    sup = PKG / "Supplementary"
    fig = sup / "Figures"
    s1 = fig / "Figure_S1.png"
    s2 = fig / "Figure_S2.png"
    out_pdf = sup / "Supporting_Information.pdf"
    tmp_pdf = sup / "_tmp_supporting_information.pdf"
    plt.rcParams["font.family"] = "DejaVu Sans"
    plt.rcParams["pdf.fonttype"] = 42
    page_counter = [0]
    from matplotlib.backends.backend_pdf import PdfPages

    with PdfPages(tmp_pdf) as pdf:
        bmc.cover_and_contents(pdf, page_counter)
        dispatch = {
            "S12": lambda: bmc.add_table_pages(pdf, "S12", page_counter),
            "F1": lambda: bmc.add_figure_page(pdf, "F1", s1, page_counter),
            "F2": lambda: bmc.add_figure_page(pdf, "F2", s2, page_counter),
            "S2": lambda: bmc.add_table_pages(pdf, "S2", page_counter),
            "S3": lambda: bmc.add_table_pages(pdf, "S3", page_counter),
            "S4": lambda: bmc.add_table_pages(pdf, "S4", page_counter),
            "S5": lambda: bmc.add_table_pages(pdf, "S5", page_counter),
            "S10": lambda: bmc.add_s10_pages(pdf, page_counter),
            "S11": lambda: bmc.add_table_pages(pdf, "S11", page_counter),
            "S7": lambda: bmc.add_table_pages(pdf, "S7", page_counter),
            "S8": lambda: bmc.add_table_pages(pdf, "S8", page_counter),
            "S9": lambda: bmc.add_table_pages(pdf, "S9", page_counter),
            "S1": lambda: bmc.add_table_pages(pdf, "S1", page_counter),
            "S6": lambda: bmc.add_table_pages(pdf, "S6", page_counter),
        }
        for k in bmc.CITE_ORDER:
            dispatch[k]()
    shutil.move(tmp_pdf, out_pdf)
    shutil.copy2(out_pdf, sup / "01_Supplementary_Information.pdf")
    for stale in (sup / "Additional_file_1.pdf",):
        if stale.exists():
            stale.unlink()

    xlsx = sup / "02_Supplementary_Tables.xlsx"
    wb = load_workbook(xlsx)
    for sheet_name in list(wb.sheetnames):
        if sheet_name.startswith("File_info") and sheet_name != "File_info":
            del wb[sheet_name]
    if "BMC_file_info" in wb.sheetnames:
        info = wb["BMC_file_info"]
        info.title = "File_info"
    elif "File_info" in wb.sheetnames:
        info = wb["File_info"]
    else:
        info = wb.create_sheet("File_info", 0)
    info["A1"] = "Supporting Information file 2"
    info["A1"].font = Font(bold=True, size=14, color="1F4E79")
    info["A2"] = "Format: XLSX (.xlsx), editable"
    info["A3"] = "Title: Complete supplementary tables S1–S12"
    info["A4"] = (
        "Description: Machine-readable tables cited in Dogan et al. (Cancer Medicine). "
        "Sheet order is numerical (S1–S12). Supporting Information file 1 (PDF) shows the same items "
        "in manuscript first-citation order, with top-20 previews of genome-wide screens."
    )
    info["A4"].alignment = Alignment(wrap_text=True)
    info["B8"] = "Short title"
    info["B8"].font = Font(bold=True)
    wb.save(xlsx)
    dest = sup / "Supporting_Information_tables.xlsx"
    shutil.copy2(xlsx, dest)
    extra = sup / "Additional_file_2.xlsx"
    if extra.exists():
        extra.unlink()

    (fig / "00_Figure_captions.txt").write_text(
        "Cancer Medicine — supporting figures (Supporting Information file 1)\n"
        "Upload the PDF separately. High-resolution PNG copies: Figure_S1.png, Figure_S2.png (300 dpi).\n\n"
        f"Figure S1. {bmc.SHORT_TITLES['F1']}.\nLegend. {bmc.LEGENDS['F1']}\n\n"
        f"Figure S2. {bmc.SHORT_TITLES['F2']}.\nLegend. {bmc.LEGENDS['F2']}\n",
        encoding="utf-8",
    )
    (sup / "Tables" / "00_Table_captions.txt").write_text(
        "Cancer Medicine — supporting tables (Supporting Information files 1 and 2)\n"
        "Tables are editable in the xlsx. Abbreviations are defined in each legend.\n"
        "Footnote symbols, when used: †, ‡, §, ¶. Asterisks are reserved for P values.\n\n"
        + "\n".join(f"Table {k}. {bmc.SHORT_TITLES[k]}." for k in [f"S{i}" for i in range(1, 13)])
        + "\n",
        encoding="utf-8",
    )
    print("SI pdf pages", page_counter[0], "MB", round(out_pdf.stat().st_size / 1e6, 2))


def build_supporting_information_versions() -> None:
    """Build a yellow review PDF, then leave clean submission PDFs in the package."""
    global REVIEW_HIGHLIGHT
    review_dir = PKG / "Supplementary" / "Review_highlighted"
    review_dir.mkdir(exist_ok=True)
    REVIEW_HIGHLIGHT = True
    build_supporting_information()
    shutil.copy2(
        PKG / "Supplementary" / "Supporting_Information.pdf",
        review_dir / "Supporting_Information_review.pdf",
    )
    REVIEW_HIGHLIGHT = False
    build_supporting_information()


def footnote_main_xlsx() -> None:
    notes = {
        "Table_1.xlsx": "†β, methylation beta value. ‡Δβ, mean β (fast-death) − mean β (other). §q, Benjamini–Hochberg false discovery rate.",
        "Table_2.xlsx": "†β, methylation beta value. ‡Δβ, mean β (fast-death) − mean β (other). §q, Benjamini–Hochberg false discovery rate.",
        "Table_3.xlsx": "†Cox β, log-hazard per unit increase in methylation beta. ‡q, Benjamini–Hochberg false discovery rate. §n, sample size. ¶Events, number of OS60 events.",
    }
    for name, note in notes.items():
        path = PKG / "Tables" / name
        wb = load_workbook(path)
        ws = wb.active
        header_row = None
        for row in ws.iter_rows(min_row=1, max_row=6):
            values = [c.value for c in row]
            if values and values[0] in ("probe", "Probe"):
                header_row = row[0].row
                break
        if header_row is None:
            raise ValueError(f"No header in {name}")
        for cell in ws[header_row]:
            if cell.value:
                cell.font = Font(bold=True, name="Times New Roman", size=10)
        last = ws.max_row + 2
        ws.cell(last, 1, note).font = Font(italic=True, name="Times New Roman", size=10)
        ws.cell(last, 1).alignment = Alignment(wrap_text=True)
        wb.save(path)


def write_captions_and_readme() -> None:
    (PKG / "Figures" / "00_Figure_captions.txt").write_text(
        "Cancer Medicine — main figures\n"
        "Upload separately, in citation order, at the highest resolution available (these PNGs are 300–600 dpi).\n"
        "Legends are in the manuscript after the tables. Each legend defines symbols and abbreviations.\n"
        "Colour figures are reproduced in colour online without charge.\n\n"
        "Figure 1. Analysis pipeline for the TCGA ESCA cohort.\n"
        "Legend. Flowchart of data sources, OS60 endpoint, penalized Cox modelling, GEO checks, fast-death methylation, and exploratory somatic layers.\n\n"
        "Figure 2. Cross-apply C-index of RNA-plus-HM450 Cox models.\n"
        "Legend. Within-lineage C-indices are apparent (in-sample). Cross-application uses fixed coefficients. Dashed line marks chance (0.5).\n\n"
        "Figure 3. Fast-death HM450 methylation contrasts by histology.\n"
        "Legend. Volcano plots (Δβ versus −log10 q_diff) and ranked hypo/hyper loci for EAC and ESCC.\n\n"
        "Figure 4. Kaplan–Meier OS60 by median probe methylation.\n"
        "Legend. High versus low HM450 beta (median split). Upper row EAC Cox-prioritised probes; lower row ESCC fast-death probes.\n\n"
        "Figure 5. Mutation-associated modules and histology interactions.\n"
        "Legend. (a) Mutant minus wild-type module-score heatmap. (b) Mutation×histology Cox interaction screen on OS60.\n\n"
        "Optional vector copy of the pipeline figure: Figure_1.pdf.\n",
        encoding="utf-8",
    )
    (PKG / "Tables" / "00_Table_captions.txt").write_text(
        "Cancer Medicine — main tables\n"
        "Editable copies of the tables in 02_Manuscript.docx. Do not supply tables as images.\n"
        "The table, legend, and footnote are readable without the main text.\n"
        "Footnote symbols: †, ‡, §, ¶. Asterisks are reserved for P values.\n\n"
        "Table 1. EAC fast-death HM450 hypo- and hypermethylated loci.\n"
        "Legend. Fast-death versus other samples in EAC/ESCA; Δβ and Benjamini–Hochberg q_diff.\n"
        "†β, methylation beta. ‡Δβ, fast-death minus other. §q, Benjamini–Hochberg false discovery rate.\n\n"
        "Table 2. ESCC fast-death HM450 hypo- and hypermethylated loci.\n"
        "Legend. Fast-death versus other samples in ESCC; Δβ and Benjamini–Hochberg q_diff.\n"
        "†β, methylation beta. ‡Δβ, fast-death minus other. §q, Benjamini–Hochberg false discovery rate.\n\n"
        "Table 3. EAC univariate Cox of probe methylation versus OS60.\n"
        "Legend. Continuous HM450 beta; cox_beta_meth is log-hazard per unit beta; cox_q_meth is BH FDR.\n"
        "†Cox β, log-hazard per unit beta. ‡q, Benjamini–Hochberg false discovery rate. §n, sample size. ¶Events, OS60 events.\n",
        encoding="utf-8",
    )
    (PKG / "00_README_upload.txt").write_text(
        "Cancer Medicine — Original Research Article upload order\n"
        "Journal: https://onlinelibrary.wiley.com/journal/20457634\n"
        "Author guidelines: https://onlinelibrary.wiley.com/page/journal/20457634/homepage/forauthors.html\n"
        "Portal: https://wiley.atyponrex.com/journal/CAM4\n"
        "Editorial office: cancermed@wiley.com\n"
        "\n"
        "Free-format submission. Required order inside the manuscript:\n"
        "title page (authors, affiliations, corresponding author, acknowledgements,\n"
        "funding statement, ethical approval, data availability, conflict of interest,\n"
        "author contributions), abstract and keywords, introduction, methods, results,\n"
        "discussion, conclusions, literature cited, tables, figure legends.\n"
        "Supporting information is uploaded as separate files.\n"
        "\n"
        "1) Cover letter\n"
        "   01_Cover_letter_Cancer_Medicine.docx\n"
        "   At least three recommended reviewers are listed in the letter.\n"
        "\n"
        "2) Manuscript (editable Word)\n"
        "   02_Manuscript.docx\n"
        "   Submitting author needs an ORCID iD in the portal.\n"
        "\n"
        "3) Main figures (citation order, highest resolution; 300–600 dpi PNG)\n"
        "   Figures/Figure_1.png … Figure_5.png\n"
        "   Optional vector: Figures/Figure_1.pdf\n"
        "\n"
        "4) Main tables\n"
        "   Already in the manuscript after the literature cited (Tables 1–3), with legends\n"
        "   and abbreviation footnotes (†, ‡, §, ¶). Asterisks are reserved for P values.\n"
        "   Optional editable Excel: Tables/Table_1.xlsx … Table_3.xlsx\n"
        "\n"
        "5) Supporting Information (separate files; not edited by production)\n"
        "   Supplementary/Supporting_Information.pdf\n"
        "      also saved as Supplementary/01_Supplementary_Information.pdf\n"
        "      Tables S1–S12 and Figures S1–S2 in first-citation order.\n"
        "   Supplementary/Supporting_Information_tables.xlsx\n"
        "      also saved as Supplementary/02_Supplementary_Tables.xlsx\n"
        "      Complete editable tables S1–S12.\n"
        "\n"
        "Do not upload: 03_Supplementary_Figures.docx (legacy);\n"
        "video_abstract/Video_Abstract_Cancer_Medicine.mp4 (on-screen title is still BMC Cancer).\n"
        "\n"
        "Keywords were left as the existing five terms. The guidelines ask for 4–6 keywords\n"
        "in the manuscript-parts list and 1–4 in the keywords subsection.\n",
        encoding="utf-8",
    )


def main() -> None:
    reformat_manuscript()
    reformat_cover_letter()
    footnote_main_xlsx()
    build_supporting_information()
    write_captions_and_readme()
    print("done")


if __name__ == "__main__":
    main()
