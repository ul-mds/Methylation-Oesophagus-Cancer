"""
Replace the single Supplementary Materials paragraph with:
  Heading 1: Supplementary Materials (unchanged)
  Intro
  Heading 2: Supplementary Tables — Supplementary Table S1..S7 with titles + result/*.tsv
  Heading 2: Supplementary Figures — unique panels only (S1–S2 mediation figures; not duplicates of main Figs 1–6)

Idempotent: skips if the first block after 'Supplementary Materials' is 'Supplementary Tables'.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

DOC_NAME = "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"

SUPP_TABLES: list[tuple[str, str]] = [
    (
        "Concordance index (C-index) for histology-specific and pooled models on overall survival.",
        "result/model_performance_cindex.tsv",
    ),
    (
        "EAC/ESCA: top hypomethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "result/methylation_fast_death_EAC_ESCA.tsv",
    ),
    (
        "ESCC: top hypomethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "result/methylation_fast_death_ESCC.tsv",
    ),
    (
        "EAC/ESCA: top hypermethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "result/methylation_fast_death_EAC_ESCA.tsv",
    ),
    (
        "ESCC: top hypermethylated HM450 probes in fast-death versus other samples (Benjamini–Hochberg FDR, q_diff).",
        "result/methylation_fast_death_ESCC.tsv",
    ),
    (
        "EAC/ESCA: univariate Cox associations between continuous probe methylation and overall survival "
        "(selected top probes by q).",
        "result/methylation_fast_death_EAC_ESCA.tsv",
    ),
    (
        "Exploratory mechanistic follow-up: strongest associations between somatic mutation and expression or module scores, "
        "stratified by histology (ranked by p-value).",
        "result/interaction_mechanism_mut_to_expr_modules.tsv",
    ),
]

SUPP_TABLE_NOTES = (
    "Tables S2–S6 are derived from the indicated methylation summary files (hypo/hyper and Cox columns). "
    "The mutation×histology interaction screen is summarized in "
    "result/mutation_x_histology_interactions_OS.tsv in addition to Table S7."
)

SUPP_FIGURES: list[tuple[str, str]] = [
    (
        "Mediation-style delta landscape (volcano): ESCC versus EAC/ESCA (supplementary to main mediation analyses).",
        "Plots/07_mediation_delta_volcano.png",
    ),
    (
        "Top histology-specific mediation contrasts (supplementary summary figure).",
        "Plots/08_top_histology_specific_mediation.png",
    ),
]


def _clear_runs(p) -> None:
    for run in list(p.runs):
        run._element.getparent().remove(run._element)


def _add_label_body(para, label: str, body: str, label_bold: bool = True) -> None:
    _clear_runs(para)
    r0 = para.add_run(label)
    r0.bold = label_bold
    if body:
        para.add_run(" " + body)


def _remove_following_paragraphs(anchor_p) -> None:
    el = anchor_p._element
    nxt = el.getnext()
    while nxt is not None:
        if nxt.tag.endswith("}sectPr"):
            break
        if not nxt.tag.endswith("}p"):
            nxt = nxt.getnext()
            continue
        kill = nxt
        nxt = kill.getnext()
        kill.getparent().remove(kill)


def _append_paragraph_after(cursor_el, doc: Document, style: str | None, text: str | None = None) -> object:
    new_el = OxmlElement("w:p")
    cursor_el.addnext(new_el)
    para = Paragraph(new_el, doc._body)
    if style:
        para.style = style
    if text is not None:
        para.add_run(text)
    return new_el


def _already_restructured(doc: Document) -> bool:
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        if p.text.strip() != "Supplementary Materials":
            continue
        for j in range(i + 1, min(i + 10, len(paras))):
            t = paras[j].text.strip()
            st = paras[j].style.name
            if t == "Supplementary Tables" and st.startswith("Heading"):
                return True
        return False
    return False


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    path = root / DOC_NAME
    doc = Document(str(path))

    supp_heading = None
    for p in doc.paragraphs:
        if p.text.strip() == "Supplementary Materials":
            supp_heading = p
            break
    if supp_heading is None:
        raise SystemExit("Could not find 'Supplementary Materials' heading.")

    if _already_restructured(doc):
        print("Supplementary Materials already restructured; skipping:", path)
        return

    _remove_following_paragraphs(supp_heading)
    cursor = supp_heading._element
    parent = doc._body

    def add(style: str | None, fn) -> None:
        nonlocal cursor
        new_el = OxmlElement("w:p")
        cursor.addnext(new_el)
        para = Paragraph(new_el, parent)
        if style:
            para.style = style
        fn(para)
        cursor = new_el

    add(
        None,
        lambda p: p.add_run(
            "Raw and processed outputs are stored alongside this manuscript under result/ (tabular data) and "
            "Plots/ (figure images). Supplementary Figures S1–S2 are mediation-focused panels not duplicated in "
            "the main article; Supplementary Tables S1–S7 support tabular reproduction as indicated below."
        ),
    )
    add("Heading 2", lambda p: p.add_run("Supplementary Tables"))

    for i, (title, relpath) in enumerate(SUPP_TABLES, start=1):
        label = f"Supplementary Table S{i}."
        add(
            None,
            lambda p, lb=label, ti=title, rp=relpath: _add_label_body(
                p,
                lb,
                f"{ti} Source file: {rp}.",
            ),
        )

    add(
        None,
        lambda p: p.add_run(SUPP_TABLE_NOTES),
    )

    add("Heading 2", lambda p: p.add_run("Supplementary Figures"))

    for i, (caption, relpath) in enumerate(SUPP_FIGURES, start=1):
        label = f"Supplementary Fig. S{i}."
        add(
            None,
            lambda p, lb=label, cap=caption, rp=relpath: _add_label_body(p, lb, f"{cap} Source file: {rp}."),
        )

    doc.save(str(path))
    print("Supplementary Materials restructured:", path)


if __name__ == "__main__":
    main()
