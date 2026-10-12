"""
Format Manuscript First.docx for Cancer Medicine submission guidelines:
https://onlinelibrary.wiley.com/page/journal/20457634/homepage/forauthors.html

- Body order: Introduction → Results → Discussion (no subheadings) → Methods
- Merge Limitations + Conclusions into Discussion
- Move Fig. 1 to after Introduction (figure order = text order)
- End matter: References → Acknowledgements → Author contributions →
  Data availability → Code availability → Competing interests →
  Additional information → Figure legends → Supplementary Information
- Nature-style reference formatting (italic journal, bold volume)
- Table Grid style; display-item note; SR checklist report
"""
from __future__ import annotations

from esca_escc_multiomic_pipeline.study_paths import study_root

import argparse
import re
import shutil
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

STUDY_ROOT = study_root()
DEFAULT_DOCX = STUDY_ROOT / "Manuscript First.docx"
W_T = qn("w:t")
W_P = qn("w:p")


def _ptext(el) -> str:
    return "".join(t.text or "" for t in el.iter(W_T)).strip()


def _para_index(doc: Document, pred) -> int:
    for i, p in enumerate(doc.paragraphs):
        if pred((p.text or "").strip(), p):
            return i
    raise ValueError(f"Paragraph not found: {pred}")


def _elements_between(doc: Document, p_start: int, p_end: int) -> list:
    body = doc.element.body
    start_el = end_el = None
    pi = -1
    for el in body:
        if el.tag == W_P:
            pi += 1
            if pi == p_start:
                start_el = el
            if pi == p_end:
                end_el = el
                break
    if start_el is None or end_el is None:
        raise ValueError(f"Range {p_start}-{p_end} not found")
    out, cap = [], False
    for el in body:
        if el is start_el:
            cap = True
        if cap:
            out.append(el)
        if el is end_el:
            break
    return out


def _elements_from(doc: Document, p_start: int) -> list:
    body = doc.element.body
    start_el = None
    pi = -1
    for el in body:
        if el.tag == W_P:
            pi += 1
            if pi == p_start:
                start_el = el
                break
    if start_el is None:
        raise ValueError(f"Start {p_start} not found")
    out, cap = [], False
    for el in body:
        if el.tag == qn("w:sectPr"):
            break
        if el is start_el:
            cap = True
        if cap:
            out.append(el)
    return out


def _rebuild_body(doc: Document, blocks: list[list]) -> None:
    body = doc.element.body
    sect_pr = next((el for el in body if el.tag == qn("w:sectPr")), None)
    for el in list(body):
        if el is not sect_pr:
            body.remove(el)
    for block in blocks:
        for el in block:
            body.append(el)
    if sect_pr is not None:
        body.append(sect_pr)


def _remove_heading_elements(block: list, headings: set[str]) -> list:
    return [el for el in block if not (el.tag == W_P and _ptext(el) in headings)]


def _collect_figure_legends(doc: Document) -> list[str]:
    legends: dict[int, str] = {}
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        m = re.match(r"^Fig\. (\d+)\.\s*(.+)", t, re.DOTALL)
        if m:
            legends[int(m.group(1))] = t
    return [legends[n] for n in sorted(legends)]


def _find_fig1_elements(doc: Document) -> list:
    """Return [image_para_el, caption_para_el] for pipeline Fig. 1."""
    cap_i = _para_index(doc, lambda t, _: t.startswith("Fig. 1.") and "Analysis pipeline" in t)
    cap_el = doc.paragraphs[cap_i]._p
    prev = cap_el.getprevious()
    out = []
    if prev is not None and prev.tag.endswith("}p"):
        if prev.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            out.append(prev)
    out.append(cap_el)
    return out


def _strip_fig1_from_block(block: list, fig1_els: set) -> list:
    return [el for el in block if id(el) not in fig1_els]


def _insert_figure_legends(doc: Document, legends: list[str]) -> None:
    """Insert Figure legends section before Supplementary Materials / Supplementary Information."""
    anchor = None
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t in ("Supplementary Materials", "Supplementary Information", "Supplementary Tables"):
            anchor = p._p
            break
    if anchor is None:
        return
    parent = doc._body
    # Skip if already present
    prev = anchor.getprevious()
    while prev is not None:
        if prev.tag.endswith("}p") and _ptext(prev) == "Figure legends":
            return
        prev = prev.getprevious()

    head = OxmlElement("w:p")
    anchor.addprevious(head)
    hp = Paragraph(head, parent)
    hp.style = doc.styles["Heading 1"]
    hp.add_run("Figure legends")

    prev_el = head
    for leg in legends:
        # Normalize Figure vs Fig. in legends block
        leg_text = re.sub(r"^Fig\. ", "Figure ", leg)
        lp = OxmlElement("w:p")
        prev_el.addnext(lp)
        para = Paragraph(lp, parent)
        m = re.match(r"^(Figure \d+\.)\s*(.*)$", leg_text, re.DOTALL)
        if m:
            r1 = para.add_run(m.group(1) + " ")
            r1.bold = True
            para.add_run(m.group(2))
        else:
            para.add_run(leg_text)
        prev_el = lp

    blank = OxmlElement("w:p")
    prev_el.addnext(blank)


def _rename_supplementary_heading(doc: Document) -> None:
    for p in doc.paragraphs:
        if (p.text or "").strip() == "Supplementary Materials":
            p.text = "Supplementary Information"
            break


def _format_reference_paragraph(p: Paragraph) -> None:
    """Apply Nature-style italic journal + bold volume via run-level edits when possible."""
    t = (p.text or "").strip()
    m = re.match(r"^(\d+)\.\s*(.+)", t)
    if not m:
        return
    num, body = m.group(1), m.group(2)
    jm = re.search(
        r"^(.+?\.\s+)("
        r"Nature|Lancet[^.]*|N Engl J Med[^.]*|"
        r"Nat\.?\s*Commun[^.]*|Nat\.?\s*Biotechnol[^.]*|"
        r"J\.?\s*Cell[^.]*|Cell Syst[^.]*|"
        r"BMC Cancer[^.]*|Cancer Cell Int[^.]*|Biomed\.?\s*Res\.?\s*Int[^.]*|"
        r"Sig\.?\s*Transduct[^.]*|Front\.?\s*[^.]+|Cancers[^.]*|"
        r"Discov\.?\s*Oncol[^.]*|J\.?\s*Gastrointest[^.]*|Protein Cell[^.]*|"
        r"ESMO Gastrointest[^.]*|Oncogene[^.]*|Int J Radiat[^.]*|"
        r"Dis Esophagus[^.]*|J\.?\s*Gastroenterol[^.]*|Nat Rev Mol Cell Biol[^.]*|"
        r"Surgery[^.]*"
        r")\s*(.*)$",
        body,
        re.DOTALL,
    )
    if not jm:
        return
    prefix, journal, tail = jm.group(1), jm.group(2), jm.group(3)
    vol_m = re.match(r"^(\d+[a-zA-Z]?)\s*,?\s*(.*)$", tail.strip())
    # Only rewrite if not already styled
    if any(r.italic for r in p.runs if (r.text or "").strip()):
        return
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.add_run(f"{num}. {prefix}")
    rj = p.add_run(journal)
    rj.italic = True
    if vol_m:
        p.add_run(" ")
        rv = p.add_run(vol_m.group(1))
        rv.bold = True
        rest = vol_m.group(2)
        if rest:
            p.add_run(", " + rest if not rest.startswith(",") else " " + rest)
    else:
        p.add_run(" " + tail)


def _format_all_references(doc: Document) -> None:
    in_refs = False
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == "References":
            in_refs = True
            continue
        if in_refs:
            if t.startswith("Acknowledgements") or t.startswith("Supplementary"):
                break
            if re.match(r"^\d+\.", t):
                _format_reference_paragraph(p)


def _format_table_captions(doc: Document) -> None:
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        m = re.match(r"^(Table \d+\.)\s*(.*)$", t, re.DOTALL)
        if not m:
            continue
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        h = p.add_run(m.group(1))
        h.bold = True
        if m.group(2):
            p.add_run(" " + m.group(2))


def _apply_table_grid(doc: Document) -> None:
    for tbl in doc.tables:
        for style in ("Table Grid", "Light Grid", "Normal Table"):
            try:
                tbl.style = style
                break
            except KeyError:
                continue


def _update_display_note(doc: Document) -> None:
    note = (
        "Cancer Medicine author guidelines note: at most eight combined figures and tables in the main article. "
        "This manuscript currently includes five figures and four tables (nine display items). "
        "Before submission, move at least one item to Supplementary Information or merge panels, "
        "and cite each in the main text using the word Supplementary "
        "(e.g. Supplementary Table S1; Supplementary Fig. S1)."
    )
    for p in doc.paragraphs:
        if "Cancer Medicine" in (p.text or "") and "display item" in (p.text or "").lower():
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            p.add_run(note)
            return


def _write_checklist_report(doc: Document, path: Path) -> None:
    paras = doc.paragraphs
    title_w = len(paras[0].text.split())
    abs_i = next(i for i, p in enumerate(paras) if p.text.strip() == "Abstract")
    abs_text = paras[abs_i + 1].text
    abs_w = len(abs_text.split())
    abs_refs = bool(re.search(r"\[\d+\]", abs_text))
    kw = next(p.text for p in paras if p.text.strip().startswith("Keywords"))
    kw_n = len([k for k in kw.replace("Keywords:", "").split(";") if k.strip()])
    intro_i = next(i for i, p in enumerate(paras) if p.text.strip() == "Introduction")
    meth_i = next(i for i, p in enumerate(paras) if p.text.strip() == "Methods")
    main = " ".join(
        p.text for p in paras[intro_i + 1 : meth_i]
        if p.text.strip() and not re.match(r"^Fig\. \d+\.", p.text.strip())
    )
    main_w = len(main.split())
    sections = []
    for name in ("Introduction", "Results", "Discussion", "Methods", "References", "Figure legends"):
        sections.append(f"  {name}: {'present' if any(p.text.strip() == name for p in paras) else 'MISSING'}")
    lines = [
        "=== Cancer Medicine format checklist (Manuscript First.docx) ===",
        f"Title words: {title_w} (limit 20) — {'PASS' if title_w <= 20 else 'FAIL'}",
        f"Abstract words: {abs_w} (limit 200) — {'PASS' if abs_w <= 200 else 'FAIL'}",
        f"Abstract no refs: {'PASS' if not abs_refs else 'FAIL'}",
        f"Keywords: {kw_n} (limit 6) — {'PASS' if kw_n <= 6 else 'FAIL'}",
        f"Main text words (Intro→Methods): {main_w} (limit 4500) — {'PASS' if main_w <= 4500 else 'FAIL'}",
        "Section order:",
        *sections,
        "Body structure: Introduction → Results → Discussion → Methods (SR guideline)",
        "End matter: References, Acknowledgements, Author contributions, Data availability, "
        "Competing interests, Figure legends, Supplementary Information",
    ]
    report = "\n".join(lines)
    out = STUDY_ROOT / "scientific_reports_format_check_report.txt"
    out.write_text(report, encoding="utf-8")
    print(report)
    print(f"\nWrote: {out}")


def format_manuscript(docx_path: Path) -> None:
    backup = docx_path.with_name(docx_path.stem + ".pre_srep.docx.bak")
    shutil.copy2(docx_path, backup)

    doc = Document(str(docx_path))
    legends = _collect_figure_legends(doc)

    # Indices (current manuscript layout)
    i_intro = _para_index(doc, lambda t, _: t == "Introduction")
    i_results = _para_index(doc, lambda t, _: t == "Results")
    i_disc = _para_index(doc, lambda t, _: t == "Discussion")
    i_lim = _para_index(doc, lambda t, _: t == "Limitations")
    i_conc = _para_index(doc, lambda t, _: t == "Conclusions")
    i_meth = _para_index(doc, lambda t, _: t == "Methods")
    i_refs = _para_index(doc, lambda t, _: t == "References")

    fig1_els = _find_fig1_elements(doc)
    fig1_ids = {id(el) for el in fig1_els}

    front = _elements_between(doc, 0, i_intro)
    intro = _elements_between(doc, i_intro, i_results - 1)
    results = _elements_between(doc, i_results, i_disc - 1)
    discussion_raw = _elements_between(doc, i_disc, i_lim - 1)
    limitations_body = _elements_between(doc, i_lim + 1, i_conc - 1)
    conclusions_body = _elements_between(doc, i_conc + 1, i_meth - 1)
    methods_raw = _elements_between(doc, i_meth, i_refs - 1)
    refs_onward = _elements_from(doc, i_refs)

    discussion = _remove_heading_elements(discussion_raw, set())
    limitations_body = _remove_heading_elements(limitations_body, {"Limitations"})
    conclusions_body = _remove_heading_elements(conclusions_body, {"Conclusions"})
    methods = _strip_fig1_from_block(methods_raw, fig1_ids)

    # SR: Introduction → Results → Discussion → Methods
    new_order = (
        front
        + intro
        + fig1_els
        + results
        + discussion
        + limitations_body
        + conclusions_body
        + methods
        + refs_onward
    )

    _rebuild_body(doc, [new_order])

    _insert_figure_legends(doc, legends)
    _rename_supplementary_heading(doc)
    _format_all_references(doc)
    _format_table_captions(doc)
    _apply_table_grid(doc)
    _update_display_note(doc)

    doc.save(str(docx_path))
    _write_checklist_report(doc, docx_path)
    print(f"\nSaved: {docx_path}")
    print(f"Backup: {backup}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", type=Path, default=DEFAULT_DOCX)
    args = ap.parse_args()
    format_manuscript(args.docx.resolve())


if __name__ == "__main__":
    main()
