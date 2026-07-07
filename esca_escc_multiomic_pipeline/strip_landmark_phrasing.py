"""
Remove repetitive '60-month landmark' / 'landmark OS' wording from manuscript text.

Applied to all body paragraphs, table cells, headers, and footers. Order matters:
longest phrases first.
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph

DOC_NAME = "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"

# (old, new) — longest matches first.
REPLACEMENTS: list[tuple[str, str]] = [
    ("Kaplan–Meier 60-month landmark overall survival", "Kaplan–Meier overall survival"),
    ("60-month landmark overall survival endpoint", "overall survival endpoint"),
    ("60-month landmark overall survival", "overall survival"),
    ("60-month landmark OS window", "60-month OS window"),
    ("60-month landmark OS", "overall survival (OS)"),
    ("(60-month landmark OS)", "(overall survival, OS)"),
    ("landmark-window deaths", "deaths within the analysis window"),
    ("among landmark-window deaths", "among deaths within the analysis window"),
    ("within the landmark window", "within the analysis window"),
    ("(landmark deaths)", "(decedents in the analysis cohort)"),
    ("landmark deaths,", "decedents in the analysis cohort,"),
    ("under a 60-month landmark overall survival endpoint", "for overall survival"),
    ("with a 60-month landmark and", "with follow-up capped at 60 months and"),
    (
        "using a 60-month landmark (follow-up capped at 60 months; deaths after 60 months censored)",
        "capping follow-up at 60 months (deaths after 60 months censored)",
    ),
    ("using a 60-month landmark ", "with 60-month capped follow-up "),
    ("a 60-month landmark overall survival endpoint", "an overall survival endpoint"),
    ("within a 60-month landmark OS window", "within a 60-month OS window"),
    ("60-month Kaplan–Meier", "Kaplan–Meier"),
    ("60-month Kaplan-Meier", "Kaplan-Meier"),
    ("to 60-month Kaplan", "to Kaplan"),
    ("stratify 60-month OS", "stratify OS"),
    ("60-month survival separation", "survival separation"),
    ("60-month survival stratification", "survival stratification"),
    ("the same landmark horizon", "the same follow-up horizon"),
    ("risk within the same landmark horizon", "risk within the same follow-up horizon"),
    ("landmark OS (nominal", "OS (nominal"),
    ("landmark OS.", "OS."),
    ("landmark OS,", "OS,"),
    ("landmark OS ", "OS "),
    ("60-month landmark omics", "60-month multi-omic analysis"),
    ("landmark survival separation", "survival separation"),
    ("landmarking can reduce", "capping follow-up can reduce"),
]


def _apply_to_text(s: str) -> str:
    t = s
    for old, new in REPLACEMENTS:
        t = t.replace(old, new)
    t = re.sub(r"  +", " ", t)
    return t


def _process_paragraph(p: Paragraph) -> bool:
    changed = False
    for run in p.runs:
        t = run.text or ""
        if not t:
            continue
        nt = _apply_to_text(t)
        if nt != t:
            run.text = nt
            changed = True
    return changed


def _flatten_paragraph_replace(p: Paragraph) -> bool:
    """Re-apply replacements on merged text (fixes phrases split across runs)."""
    full = "".join(r.text or "" for r in p.runs)
    if not full.strip() or "landmark" not in full.lower():
        return False
    if full.strip() == "OS_60mo_landmark" or "OS_60mo_landmark" in full and len(full) < 40:
        return False
    nt = _apply_to_text(full)
    if nt == full:
        return False
    for run in list(p.runs):
        run._element.getparent().remove(run._element)
    p.add_run(nt)
    return True


def _remove_orphan_duplicate_fig6_legend(doc: Document) -> bool:
    """Remove a full Figure 6 legend paragraph if it sits between tables and Discussion (duplicate of end-matter legend)."""
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        if (
            t.startswith("Fig. 6 Mutation×histology interaction screen on overall survival (OS)")
            and "Exploratory Cox screen for somatic mutation×histology interaction" in t
            and "Quasi-complete separation cases omitted" in t
        ):
            p._element.getparent().remove(p._element)
            return True
    return False


def _rename_os_landmark_column_in_tables(doc: Document) -> int:
    n = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text and "OS_60mo_landmark" in r.text:
                            r.text = r.text.replace("OS_60mo_landmark", "OS_60mo")
                            n += 1
    return n


def strip_document(doc: Document) -> int:
    n = 0
    for p in doc.paragraphs:
        if _process_paragraph(p):
            n += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _process_paragraph(p):
                        n += 1
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                if _process_paragraph(p):
                    n += 1
    for p in doc.paragraphs:
        if _flatten_paragraph_replace(p):
            n += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _flatten_paragraph_replace(p):
                        n += 1
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                if _flatten_paragraph_replace(p):
                    n += 1
    return n


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    path = root / DOC_NAME
    doc = Document(str(path))
    changed = strip_document(doc)
    if _remove_orphan_duplicate_fig6_legend(doc):
        changed += 1
        print("Removed orphan duplicate Figure 6 legend paragraph.")
    tbl = _rename_os_landmark_column_in_tables(doc)
    if tbl:
        changed += tbl
        print(f"Renamed OS_60mo_landmark -> OS_60mo in {tbl} table cell run(s).")
    doc.save(str(path))
    print(f"Landmark phrasing pass: updated {changed} paragraph/run/table edits. Saved {path}")


if __name__ == "__main__":
    main()
