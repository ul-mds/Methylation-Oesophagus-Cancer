"""Build BMC Cancer Additional files for the ESCA submission package.

Additional file 1 (PDF): supplementary tables and figures in first-citation order.
Additional file 2 (XLSX): complete machine-readable tables S1–S12.
"""
from __future__ import annotations

import math
import shutil
import textwrap
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, PatternFill as PF
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

ROOT = Path("/Users/senol/Desktop/AML/Esophageal Carcinoma (TCGA, Firehose Legacy)")
PKG = ROOT / "Cancer Medicine submission_package"
SUP = PKG / "Supplementary"
XLSX_SRC = SUP / "02_Supplementary_Tables.xlsx"
TAB = SUP / "Tables"
FIG = SUP / "Figures"

A4 = (8.27, 11.69)
MARGIN = 0.55

SHORT_TITLES = {
    "S1": "C-index for histology-specific and pooled Cox models",
    "S2": "EAC short-survival-associated hypomethylated gene-level HM450 features",
    "S3": "ESCC short-survival-associated hypomethylated gene-level HM450 features",
    "S4": "EAC short-survival-associated hypermethylated gene-level HM450 features",
    "S5": "ESCC short-survival-associated hypermethylated gene-level HM450 features",
    "S6": "EAC univariate methylation–OS60 Cox screen",
    "S7": "Mutation-associated expression and module shifts",
    "S8": "GSE161533 overlap with pre-specified ESCC genes",
    "S9": "GSE53624/GSE53622 lncRNA OS60 Cox transport",
    "S10": "Short-survival threshold sensitivity (10/20/30%)",
    "S11": "Mutation-by-histology interaction screen on OS60",
    "S12": "REMARK checklist for prognostic model reporting",
    "F1": "Mediation-style methylation–expression delta volcano",
    "F2": "Top histology-specific mediation contrasts",
}

LEGENDS = {
    "S1": (
        "Harrell C-index on OS60 for penalized Cox models trained on paired RNA z-scores and HM450 "
        "beta values. Model indicates the training histology; test indicates the evaluation set. "
        "Within-lineage values are apparent (in-sample) estimates. n is the evaluation sample size. "
        "Complete numeric matrix corresponding to Fig. 2."
    ),
    "S2": (
        "Gene-level HumanMethylation450 (HM450) features ranked among esophageal adenocarcinoma "
        "(EAC) samples for lower methylation in the short-survival group "
        "(delta = mean beta_short_survival − mean beta_other < 0). q_diff is the "
        "Benjamini–Hochberg false discovery rate (FDR). The PDF shows the top 20 features; "
        "the complete ranked list is in Supporting Information file 2."
    ),
    "S3": "As Table S2, for esophageal squamous cell carcinoma (ESCC) hypomethylation. PDF: top 20 features; complete list in Supporting Information file 2.",
    "S4": "As Table S2, for EAC hypermethylation (delta > 0). PDF: top 20 features; complete list in Supporting Information file 2.",
    "S5": "As Table S3, for ESCC hypermethylation (delta > 0). PDF: top 20 features; complete list in Supporting Information file 2.",
    "S6": (
        "Univariate Cox models of continuous gene-level HM450 beta versus overall survival "
        "truncated at 60 months (OS60) in EAC. cox_beta_meth is the log-hazard per unit beta; "
        "cox_q_meth is the Benjamini–Hochberg FDR. PDF: top 20 by cox_q_meth; complete screen "
        "in Supporting Information file 2."
    ),
    "S7": (
        "Exploratory mutant versus wild-type contrasts on RNA z-scores and Hallmark module scores, "
        "stratified by histology. p values are nominal two-sided Welch tests. "
        "PDF: 20 lowest p values; full table in Additional file 2."
    ),
    "S8": (
        "Pre-specified TCGA ESCC gene union (n = 110) overlapping gene-level summaries of the top "
        "1200 GSE161533 paired tumour–normal probes (n = 28 patients)."
    ),
    "S9": (
        "60-month truncated elastic-net Cox on high-variance lncRNA features: train GSE53624, "
        "hold-out score GSE53622. This module does not replicate TCGA HM450 short-survival methylation."
    ),
    "S10": (
        "Short survival was defined as the shortest-survival tail among OS60 decedents (minimum n = 10) at "
        "10%, 20% and 30% thresholds. Panel A: cohort counts and Jaccard overlap. Panel B: key loci "
        "ranks including SHISA3 (ESCC) and SIGLEC14/EPHB1/ZNF214/RAB11FIP3 (EAC)."
    ),
    "S11": (
        "Cox models of mutation + histology + mutation×histology on OS60 for the 80 most frequent "
        "genes with ≥8 mutants. p_interaction is the nominal Wald p for the product term (not "
        "genome-wide adjusted). PDF: 20 smallest p_interaction; full screen in Additional file 2."
    ),
    "S12": (
        "REMARK items for this secondary TCGA analysis (McShane et al., J Clin Oncol 2005). "
        "Y = reported; P = partial; N = not applicable."
    ),
    "F1": (
        "Volcano plot of the difference in mediation-style indirect-effect proxies between "
        "esophageal adenocarcinoma (EAC) and esophageal squamous cell carcinoma (ESCC). The x-axis "
        "shows EAC minus ESCC; the y-axis shows −log10 of the smaller expression-coefficient P value. "
        "Full overall survival (OS) was used for exploratory feature prioritization. This is a "
        "hypothesis-generating analysis, not an OS60 validation analysis."
    ),
    "F2": (
        "Ranked differences in mediation-style indirect-effect proxies (EAC minus ESCC) using full "
        "overall survival (OS) for exploratory feature prioritization. Positive values are larger in "
        "EAC and negative values are larger in ESCC. This is not an OS60 validation analysis or a "
        "clinical signature."
    ),
}

# First-citation order in 02_Manuscript.docx (Methods → Results → Discussion).
CITE_ORDER = [
    "S12", "F1", "F2", "S2", "S3", "S4", "S5", "S10", "S11", "S7", "S8", "S9", "S1", "S6",
]


def wrap(s: str, width: int) -> str:
    return "\n".join(textwrap.wrap(s, width=width, break_long_words=False) or [s])


def new_page(pdf: PdfPages, title: str | None = None):
    fig = plt.figure(figsize=A4)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.text(MARGIN / A4[0], 1 - 0.32 / A4[1], "Additional file 1  ·  BMC Cancer", fontsize=8, color="#444444")
    ax.text(1 - MARGIN / A4[0], 1 - 0.32 / A4[1], "Dogan et al.", fontsize=8, color="#444444", ha="right")
    ax.plot([MARGIN / A4[0], 1 - MARGIN / A4[0]], [1 - 0.42 / A4[1], 1 - 0.42 / A4[1]], color="#1f4e79", lw=0.8)
    if title:
        ax.text(MARGIN / A4[0], 1 - 0.72 / A4[1], title, fontsize=12, fontweight="bold", va="top", color="#1f4e79")
    return fig, ax


def footer(ax, page: int, n_pages: int | None = None):
    lab = f"Page {page}" if n_pages is None else f"Page {page} of {n_pages}"
    ax.text(0.5, 0.28 / A4[1], lab, ha="center", fontsize=8, color="#666666")


def fmt_num(x):
    if pd.isna(x):
        return ""
    if isinstance(x, (int, np.integer)):
        return str(int(x))
    if isinstance(x, (float, np.floating)):
        if abs(x) >= 1000 or x == 0:
            return f"{x:.4g}"
        if abs(x) >= 1:
            return f"{x:.4f}".rstrip("0").rstrip(".")
        return f"{x:.3g}"
    s = str(x)
    return s if len(s) <= 28 else s[:25] + "…"


def load_preview(key: str) -> pd.DataFrame:
    xlsx = XLSX_SRC
    if key == "S1":
        df = pd.read_excel(xlsx, sheet_name="S1")
    elif key == "S2":
        df = pd.read_csv(TAB / "S02_methylation_fast_death_EAC_ESCA.tsv", sep="\t")
        df = df.sort_values("q_diff").head(20)
        df = df[["gene", "mean_beta_short_survival", "mean_beta_other", "delta_short_survival_minus_other", "q_diff"]]
    elif key == "S3":
        df = pd.read_csv(TAB / "S03_methylation_fast_death_ESCC.tsv", sep="\t")
        df = df.sort_values("q_diff").head(20)
        df = df[["gene", "mean_beta_short_survival", "mean_beta_other", "delta_short_survival_minus_other", "q_diff"]]
    elif key == "S4":
        df = pd.read_csv(TAB / "S04_methylation_fast_death_EAC_ESCA.tsv", sep="\t")
        df = df.sort_values("q_diff").head(20)
        df = df[["gene", "mean_beta_short_survival", "mean_beta_other", "delta_short_survival_minus_other", "q_diff"]]
    elif key == "S5":
        df = pd.read_csv(TAB / "S05_methylation_fast_death_ESCC.tsv", sep="\t")
        df = df.sort_values("q_diff").head(20)
        df = df[["gene", "mean_beta_short_survival", "mean_beta_other", "delta_short_survival_minus_other", "q_diff"]]
    elif key == "S6":
        df = pd.read_csv(TAB / "S06_methylation_OS60_Cox_EAC_ESCA.tsv", sep="\t")
        df = df.sort_values("cox_q_meth").head(20)
        df = df[["gene", "cox_beta_meth", "cox_q_meth", "cox_n", "cox_events"]]
    elif key == "S7":
        df = pd.read_excel(xlsx, sheet_name="S7").sort_values("p").head(20)
        df = df[["gene", "histology", "feature", "mut_n", "wt_n", "delta", "p"]]
    elif key == "S8":
        df = pd.read_excel(xlsx, sheet_name="S8")
    elif key == "S9":
        df = pd.read_excel(xlsx, sheet_name="S9")
    elif key == "S10":
        df = pd.read_excel(xlsx, sheet_name="S10_summary")
    elif key == "S11":
        df = pd.read_excel(xlsx, sheet_name="S11").sort_values("p_interaction").head(20)
        df = df[["gene", "n", "events", "mutated_n", "beta_interaction", "p_interaction", "q_interaction"]]
    elif key == "S12":
        df = pd.read_excel(xlsx, sheet_name="S12")
        # align funding status with manuscript Declarations
        df.loc[df["Item"] == 25, "Status"] = "Y"
        df.loc[df["Item"] == 25, "Where reported"] = "Declarations — Tag-White grants; Leipzig OA fund"
    else:
        raise KeyError(key)
    return df.reset_index(drop=True)


def draw_table_block(ax, df: pd.DataFrame, y0: float) -> float:
    show = df.copy()
    show.columns = [str(c) for c in show.columns]
    for c in show.columns:
        show[c] = show[c].map(fmt_num)
    col_labels = [c.replace("_", "\n") for c in show.columns]
    cell = show.values.tolist()
    nrows, ncols = show.shape
    table = ax.table(
        cellText=cell,
        colLabels=col_labels,
        loc="upper left",
        bbox=[MARGIN / A4[0], 0.10, 1 - 2 * MARGIN / A4[0], y0 - 0.14],
    )
    table.auto_set_font_size(False)
    fs = 6.2 if ncols <= 7 else 5.4
    table.set_fontsize(fs)
    for (r, c), cell_obj in table.get_celld().items():
        cell_obj.set_edgecolor("#cccccc")
        cell_obj.set_linewidth(0.3)
        if r == 0:
            cell_obj.set_facecolor("#1f4e79")
            cell_obj.set_text_props(color="white", fontweight="bold", fontsize=fs)
        elif r % 2 == 0:
            cell_obj.set_facecolor("#f4f7fb")
        cell_obj.PAD = 0.02
    return 0.10


def add_table_pages(pdf: PdfPages, key: str, page_counter: list[int]):
    df = load_preview(key)
    title = f"Additional file 1 — Table {key}. {SHORT_TITLES[key]}"
    legend = LEGENDS[key]
    # paginate wide/long tables
    max_rows = 22 if key != "S10" else 16
    n_chunks = max(1, math.ceil(len(df) / max_rows))
    for i in range(n_chunks):
        chunk = df.iloc[i * max_rows : (i + 1) * max_rows]
        fig, ax = new_page(pdf, title if i == 0 else title + f" (continued {i + 1}/{n_chunks})")
        page_counter[0] += 1
        y = 0.88
        if i == 0:
            ax.text(
                MARGIN / A4[0],
                y,
                wrap("Legend. " + legend, 108),
                fontsize=8,
                va="top",
                wrap=False,
            )
            y = 0.88 - 0.018 * (wrap("Legend. " + legend, 108).count("\n") + 2)
        ax.text(
            MARGIN / A4[0],
            y,
            f"Displaying {len(chunk)} of {len(df) if key not in ('S2','S3','S4','S5','S6','S7','S11') else 'full Additional file 2'} rows"
            if key in ("S2", "S3", "S4", "S5", "S6", "S7", "S11")
            else f"Rows {i * max_rows + 1}–{i * max_rows + len(chunk)} of {len(df)}.",
            fontsize=7.5,
            color="#333333",
            va="top",
        )
        y -= 0.03
        draw_table_block(ax, chunk, y)
        footer(ax, page_counter[0])
        pdf.savefig(fig)
        plt.close(fig)


def add_s10_pages(pdf: PdfPages, page_counter: list[int]):
    title = f"Additional file 1 — Table S10. {SHORT_TITLES['S10']}"
    legend = LEGENDS["S10"]
    parts = [
        ("Panel A. Cohort membership and Jaccard overlap.", pd.read_excel(XLSX_SRC, sheet_name="S10_summary")),
        ("Panel B. Key loci across 10/20/30% short-survival thresholds.", pd.read_excel(XLSX_SRC, sheet_name="S10_key_loci")),
    ]
    for pi, (subtitle, df) in enumerate(parts):
        max_rows = 18
        n_chunks = max(1, math.ceil(len(df) / max_rows))
        for i in range(n_chunks):
            chunk = df.iloc[i * max_rows : (i + 1) * max_rows]
            head = title if pi == 0 and i == 0 else title + " (continued)"
            fig, ax = new_page(pdf, head)
            page_counter[0] += 1
            y = 0.88
            if pi == 0 and i == 0:
                ax.text(MARGIN / A4[0], y, wrap("Legend. " + legend, 108), fontsize=8, va="top")
                y -= 0.018 * (wrap("Legend. " + legend, 108).count("\n") + 2)
            ax.text(MARGIN / A4[0], y, subtitle, fontsize=9, fontweight="bold", va="top", color="#1f4e79")
            y -= 0.035
            draw_table_block(ax, chunk, y)
            footer(ax, page_counter[0])
            pdf.savefig(fig)
            plt.close(fig)


def add_figure_page(pdf: PdfPages, key: str, img_path: Path, page_counter: list[int]):
    title = f"Additional file 1 — Figure {key.replace('F', 'S')}. {SHORT_TITLES[key]}"
    fig, ax = new_page(pdf, title)
    page_counter[0] += 1
    ax.text(MARGIN / A4[0], 0.88, wrap("Legend. " + LEGENDS[key], 108), fontsize=8, va="top")
    im = PILImage.open(img_path).convert("RGB")
    left = MARGIN / A4[0]
    right = 1 - MARGIN / A4[0]
    bottom = 0.08
    top = 0.78
    box_w = (right - left) * A4[0]
    box_h = (top - bottom) * A4[1]
    iw, ih = im.size
    aspect = iw / ih
    box_aspect = box_w / box_h
    if aspect > box_aspect:
        w = right - left
        h = w * A4[0] / aspect / A4[1]
        x0 = left
        y0 = bottom + ((top - bottom) - h) / 2
    else:
        h = top - bottom
        w = h * A4[1] * aspect / A4[0]
        x0 = left + ((right - left) - w) / 2
        y0 = bottom
    ax_im = fig.add_axes([x0, y0, w, h])
    ax_im.imshow(im)
    ax_im.axis("off")
    footer(ax, page_counter[0])
    pdf.savefig(fig, dpi=300)
    plt.close(fig)


def cover_and_contents(pdf: PdfPages, page_counter: list[int]):
    fig, ax = new_page(pdf)
    page_counter[0] += 1
    ax.text(MARGIN / A4[0], 0.90, "Additional file 1", fontsize=18, fontweight="bold", color="#1f4e79")
    ax.text(
        MARGIN / A4[0],
        0.85,
        wrap(
            "Histology-specific prognostic architecture in TCGA esophageal carcinoma: "
            "cross-histology transport failure and lineage-specific methylation patterns",
            88,
        ),
        fontsize=11,
        va="top",
    )
    ax.text(
        MARGIN / A4[0],
        0.74,
        wrap(
            "Senol Dogan, Rene Thieme, Sanne Hoefnagel, Chuang Yang, Toralf Kirsten. "
            "BMC Cancer Research article — supplementary tables and figures.",
            100,
        ),
        fontsize=9,
        va="top",
    )
    ax.text(
        MARGIN / A4[0],
        0.66,
        wrap(
            "This file presents every supplementary table (S1–S12) and supplementary figure (S1–S2) "
            "in the order they are first cited in the manuscript. Genome-wide methylation screens "
            "(S2–S6) are shown as the top 20 ranked loci; complete matrices are in Additional file 2 "
            "(.xlsx). Table titles are ≤15 words; legends are ≤300 words, per BMC Cancer table/figure rules.",
            108,
        ),
        fontsize=9,
        va="top",
    )
    ax.text(MARGIN / A4[0], 0.52, "Contents (first-citation order)", fontsize=12, fontweight="bold", color="#1f4e79")
    labels = {
        "S12": "Table S12 — REMARK checklist",
        "F1": "Figure S1 — Mediation-style delta volcano",
        "F2": "Figure S2 — Top histology-specific mediation contrasts",
        "S2": "Table S2 — EAC short-survival hypomethylation (top 20)",
        "S3": "Table S3 — ESCC short-survival hypomethylation (top 20)",
        "S4": "Table S4 — EAC short-survival hypermethylation (top 20)",
        "S5": "Table S5 — ESCC short-survival hypermethylation (top 20)",
        "S10": "Table S10 — Short-survival 10/20/30% sensitivity",
        "S11": "Table S11 — Mutation×histology interaction screen (top 20)",
        "S7": "Table S7 — Mutation–expression/module shifts (top 20)",
        "S8": "Table S8 — GSE161533 gene overlap",
        "S9": "Table S9 — GSE53624/GSE53622 lncRNA Cox",
        "S1": "Table S1 — Cross-apply C-index matrix",
        "S6": "Table S6 — EAC methylation–OS60 Cox (top 20)",
    }
    y = 0.48
    for i, k in enumerate(CITE_ORDER, 1):
        ax.text(MARGIN / A4[0], y, f"{i:02d}.  {labels[k]}", fontsize=9, family="DejaVu Sans", va="top")
        y -= 0.026
    ax.text(
        MARGIN / A4[0],
        0.08,
        "Numerical order S1–S12 then Figures S1–S2 is preserved in Additional file 2 sheet order.",
        fontsize=8,
        color="#444444",
    )
    footer(ax, page_counter[0])
    pdf.savefig(fig)
    plt.close(fig)


def write_s12_tsv():
    df = load_preview("S12")
    out = TAB / "S12_REMARK_checklist.tsv"
    df.to_csv(out, sep="\t", index=False)
    return df


def patch_xlsx_s12_funding():
    wb = load_workbook(XLSX_SRC)
    ws = wb["S12"]
    # header row 1; items 1–25 in subsequent rows
    for row in ws.iter_rows(min_row=2, max_col=4):
        item = row[0].value
        if item == 25:
            row[2].value = "Y"
            row[3].value = "Declarations — Tag-White grants; Leipzig OA fund"
    # BMC cover sheet
    if "BMC_file_info" in wb.sheetnames:
        del wb["BMC_file_info"]
    ws0 = wb.create_sheet("BMC_file_info", 0)
    ws0["A1"] = "Additional file 2"
    ws0["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws0["A2"] = "Format: XLSX (.xlsx)"
    ws0["A3"] = "Title: Complete supplementary tables S1–S12"
    ws0["A4"] = (
        "Description: Machine-readable tables cited in Dogan et al. (BMC Cancer). "
        "Sheet order is numerical (S1–S12). Additional file 1 (PDF) shows the same items "
        "in manuscript first-citation order, with top-20 previews of genome-wide screens."
    )
    ws0["A4"].alignment = Alignment(wrap_text=True)
    ws0.merge_cells("A4:F6")
    ws0["A8"] = "Sheet"
    ws0["B8"] = "BMC short title (≤15 words)"
    ws0["A8"].font = Font(bold=True)
    ws0["B8"].font = Font(bold=True)
    for i, k in enumerate(["S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8", "S9", "S10_summary", "S10_key_loci", "S11", "S12"], start=9):
        ws0[f"A{i}"] = k
        key = "S10" if k.startswith("S10") else k
        ws0[f"B{i}"] = SHORT_TITLES[key]
    ws0.column_dimensions["A"].width = 16
    ws0.column_dimensions["B"].width = 62
    wb.save(XLSX_SRC)


def update_manuscript_additional_files():
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX

    path = PKG / "02_Manuscript.docx"
    doc = Document(str(path))
    # Replace the Supplementary Information block body with BMC Additional files list
    si = None
    for p in doc.paragraphs:
        if p.text.strip() == "Supplementary Information":
            si = p
            break
    if si is None:
        return
    # yellow highlight heading rename
    si_run_done = False
    if si.runs:
        si.runs[0].text = "Additional files"
        si.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        for r in si.runs[1:]:
            r.text = ""
        si_run_done = True
    # rewrite following non-heading paragraphs until next empty end: keep captions but prefix Additional file mapping
    new_intro = (
        "Supplementary tables and figures are provided as Additional files and are cited in the text as "
        "Tables S1–S12 and Figs. S1–S2. Additional file 1 presents every item in first-citation order; "
        "Additional file 2 contains the complete spreadsheets."
    )
    # first normal paragraph after heading
    nxt = si._element.getnext()
    from docx.text.paragraph import Paragraph

    if nxt is not None:
        p1 = Paragraph(nxt, si._parent)
        # clear and set
        for r in list(p1.runs):
            r.text = ""
        if p1.runs:
            p1.runs[0].text = new_intro
            p1.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            run = p1.add_run(new_intro)
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    # append BMC file list before figure legends if not present
    texts = [p.text for p in doc.paragraphs]
    if not any(t.startswith("Additional file 1.") for t in texts):
        # insert before Figure legends
        fl = None
        for p in doc.paragraphs:
            if p.text.strip() == "Figure legends":
                fl = p
                break
        if fl is not None:
            def ins_before(paragraph, text, heading=False):
                new_el = paragraph._p.makeelement(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p", {}
                )
                paragraph._p.addprevious(new_el)
                np = Paragraph(new_el, paragraph._parent)
                if heading:
                    np.style = "Heading 1"
                run = np.add_run(text)
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                return np

            ins_before(
                fl,
                "Additional file 2. File format: XLSX (.xlsx). Title: Complete supplementary tables S1–S12. "
                "Description: Machine-readable tables corresponding to Tables S1–S12, including full genome-wide "
                "HM450 ranks (S2–S6), mutation screens (S7, S11), GEO modules (S8–S9), short-survival sensitivity (S10) "
                "and the REMARK checklist (S12).",
            )
            ins_before(
                fl,
                "Additional file 1. File format: PDF (.pdf). Title: Supplementary tables and figures in citation order. "
                "Description: Tables S1–S12 and Figures S1–S2 laid out in the order first cited in the manuscript, "
                "with BMC short titles and legends. Genome-wide tables show the top 20 ranked loci.",
            )
    doc.save(str(path))


def main():
    SUP.mkdir(parents=True, exist_ok=True)
    (SUP / "Figures").mkdir(exist_ok=True)
    (SUP / "Tables").mkdir(exist_ok=True)

    write_s12_tsv()
    patch_xlsx_s12_funding()

    # BMC figure filenames
    s1_src = FIG / "Supplementary_Figure_S1.png"
    s2_src = FIG / "Supplementary_Figure_S2.png"
    s1 = FIG / "Figure_S1.png"
    s2 = FIG / "Figure_S2.png"
    if s1_src.exists():
        shutil.copy2(s1_src, s1)
    if s2_src.exists():
        shutil.copy2(s2_src, s2)

    out_pdf = SUP / "01_Supplementary_Information.pdf"
    tmp_pdf = SUP / "_tmp_additional_file_1.pdf"
    plt.rcParams["font.family"] = "DejaVu Sans"
    plt.rcParams["pdf.fonttype"] = 42

    page_counter = [0]
    with PdfPages(tmp_pdf) as pdf:
        cover_and_contents(pdf, page_counter)
        dispatch = {
            "S12": lambda: add_table_pages(pdf, "S12", page_counter),
            "F1": lambda: add_figure_page(pdf, "F1", s1, page_counter),
            "F2": lambda: add_figure_page(pdf, "F2", s2, page_counter),
            "S2": lambda: add_table_pages(pdf, "S2", page_counter),
            "S3": lambda: add_table_pages(pdf, "S3", page_counter),
            "S4": lambda: add_table_pages(pdf, "S4", page_counter),
            "S5": lambda: add_table_pages(pdf, "S5", page_counter),
            "S10": lambda: add_s10_pages(pdf, page_counter),
            "S11": lambda: add_table_pages(pdf, "S11", page_counter),
            "S7": lambda: add_table_pages(pdf, "S7", page_counter),
            "S8": lambda: add_table_pages(pdf, "S8", page_counter),
            "S9": lambda: add_table_pages(pdf, "S9", page_counter),
            "S1": lambda: add_table_pages(pdf, "S1", page_counter),
            "S6": lambda: add_table_pages(pdf, "S6", page_counter),
        }
        for k in CITE_ORDER:
            dispatch[k]()

    shutil.move(tmp_pdf, out_pdf)
    shutil.copy2(out_pdf, SUP / "Additional_file_1.pdf")
    shutil.copy2(XLSX_SRC, SUP / "Additional_file_2.xlsx")

    # BMC captions for supplementary figures
    (FIG / "00_Figure_captions.txt").write_text(
        "BMC Cancer — supplementary figures (Additional file 1)\n\n"
        f"Figure S1. {SHORT_TITLES['F1']}.\nLegend. {LEGENDS['F1']}\n\n"
        f"Figure S2. {SHORT_TITLES['F2']}.\nLegend. {LEGENDS['F2']}\n",
        encoding="utf-8",
    )
    (TAB / "00_Table_captions.txt").write_text(
        "BMC Cancer — supplementary tables (Additional files 1 and 2)\n"
        "Short titles ≤15 words.\n\n"
        + "\n".join(f"Table {k}. {SHORT_TITLES[k]}." for k in [f"S{i}" for i in range(1, 13)])
        + "\n",
        encoding="utf-8",
    )

    update_manuscript_additional_files()

    size_mb = out_pdf.stat().st_size / 1e6
    print("PDF pages (approx counter)", page_counter[0], "size_MB", round(size_mb, 2))
    print("Wrote", out_pdf)
    print("Wrote", SUP / "Additional_file_1.pdf")
    print("Wrote", SUP / "Additional_file_2.xlsx")


if __name__ == "__main__":
    main()
