"""
Audit and fix Manuscript First.docx for Scientific Reports:
- Nature-style reference list (italic journal, bold volume, et al. for 6+ authors)
- Supplementary Table/Fig. citation phrasing per SR guidelines
- Remove duplicate figure-legend heading
- Embed Supplementary Fig. S1–S2 PNGs above captions
"""
from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from lxml import etree

STUDY_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DOCX = STUDY_ROOT / "Manuscript First.docx"
PLOTS = STUDY_ROOT / "Plots"

from fix_verified_references_manuscript import VERIFIED_REFS as NATURE_REFS

SUPP_REPLACEMENTS: list[tuple[str, str]] = [
    (
        "Supplementary Fig.s S1–S2",
        "Supplementary Figs. S1–S2",
    ),
    (
        "with Supplementary Table S7 and Supplementary Figs. S1–S2 providing extended rankings and mediation-style contrasts to support future functional validation (Supplementary Figs. S1–S2).",
        "with extended rankings in Supplementary Table S7 online and mediation-style contrasts in Supplementary Figs. S1–S2 online to support future functional validation.",
    ),
    (
        "(with complete ranked exports in Supplementary Tables S2–S5 (ESCA/EAC: S2 for Δβ<0, S4 for Δβ>0; ESCC: S3 for Δβ<0, S5 for Δβ>0)).",
        "(complete ranked exports in Supplementary Tables S2–S5 online).",
    ),
    (
        "Supplementary Tables S1–S7 and Supplementary Figs. S1–S2 ",
        "Supplementary Tables S1–S12 and Supplementary Figs. S1–S2 ",
    ),
    (
        "Supplementary Tables S1–S7 and Supplementary Figs. S1–S2 catalogue",
        "Supplementary Tables S1–S12 and Supplementary Figs. S1–S2 catalogue",
    ),
    (
        "Supplementary Tables S1–S7 and Supplementary Figs. S1–S2 provide",
        "Supplementary Tables S1–S12 and Supplementary Figs. S1–S2 provide",
    ),
    (
        "Supplementary Tables S1–S7 and Supplementary Figs. S1–S2 are",
        "Supplementary Tables S1–S12 and Supplementary Figs. S1–S2 are",
    ),
    (
        "Supplementary Tables S1–S7 and Supplementary Figs. S1–S2 support",
        "Supplementary Tables S1–S12 and Supplementary Figs. S1–S2 support",
    ),
]

SUPP_FIG_SPECS: list[tuple[str, str, float]] = [
    ("Supplementary Fig. S1.", "07_mediation_delta_volcano.png", 6.2),
    ("Supplementary Fig. S2.", "08_top_histology_specific_mediation.png", 6.2),
]


def _replace_in_paragraph(p: Paragraph, replacements: list[tuple[str, str]]) -> None:
    from docx.oxml.ns import qn

    w_t = qn("w:t")
    for t_el in p._p.iter(w_t):
        if not t_el.text:
            continue
        new = t_el.text
        for old, rep in replacements:
            if old in new:
                new = new.replace(old, rep)
        if new != t_el.text:
            t_el.text = new


def _replace_in_all(doc: Document, replacements: list[tuple[str, str]]) -> None:
    for p in doc.paragraphs:
        _replace_in_paragraph(p, replacements)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, replacements)


def _set_paragraph_text(p: Paragraph, new_text: str) -> None:
    from docx.oxml.ns import qn

    w_t = qn("w:t")
    ts = list(p._p.iter(w_t))
    if not ts:
        p.add_run(new_text)
        return
    ts[0].text = new_text
    for t in ts[1:]:
        t.text = ""


def _fix_supplementary_paragraphs(doc: Document) -> None:
    """Whole-paragraph fixes when citations span multiple Word runs."""
    for p in doc.paragraphs:
        t = p.text or ""
        if "Supplementary Tables S2–S5 (ESCA/EAC:" in t:
            _set_paragraph_text(
                p,
                t.replace(
                    "(with complete ranked exports in Supplementary Tables S2–S5 (ESCA/EAC: S2 for Δβ<0, S4 for Δβ>0; ESCC: S3 for Δβ<0, S5 for Δβ>0)).",
                    "(complete ranked exports in Supplementary Tables S2–S5 online).",
                ),
            )
        if "Supplementary Fig.s" in t or (
            "Supplementary Table S7 and Supplementary Figs. S1–S2 providing" in t
        ):
            new = t.replace("Supplementary Fig.s S1–S2", "Supplementary Figs. S1–S2")
            new = new.replace(
                " with Supplementary Table S7 and Supplementary Figs. S1–S2 providing extended rankings and mediation-style contrasts to support future functional validation (Supplementary Figs. S1–S2).",
                " with extended rankings in Supplementary Table S7 online and mediation-style contrasts in Supplementary Figs. S1–S2 online to support future functional validation.",
            )
            if new != t:
                _set_paragraph_text(p, new)


def _write_nature_ref(p: Paragraph, num: int, prefix: str, journal: str, volume: str | None, suffix: str) -> None:
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.add_run(f"{num}. {prefix}")
    rj = p.add_run(journal)
    rj.italic = True
    if volume:
        p.add_run(" ")
        rv = p.add_run(volume)
        rv.bold = True
    p.add_run(suffix)


def _rewrite_reference_list(doc: Document) -> None:
    in_refs = False
    idx = 0
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == "References":
            in_refs = True
            continue
        if in_refs:
            if t.startswith("Acknowledgements"):
                break
            if re.match(r"^\d+\.", t):
                if idx >= len(NATURE_REFS):
                    raise RuntimeError(f"More reference paragraphs than NATURE_REFS ({idx + 1})")
                prefix, journal, volume, suffix = NATURE_REFS[idx]
                _write_nature_ref(p, idx + 1, prefix, journal, volume, suffix)
                idx += 1
    if idx != len(NATURE_REFS):
        raise RuntimeError(f"Expected {len(NATURE_REFS)} references, rewrote {idx}")


def _remove_duplicate_legend_heading(doc: Document) -> None:
    seen = False
    to_remove = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == "Figure legends":
            if seen:
                to_remove.append(p._p)
            seen = True
        elif t == "Fig. legends":
            to_remove.append(p._p)
    for el in to_remove:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)


def _normalize_supplementary_captions(doc: Document) -> None:
    """Ensure end-matter labels match SR: Supplementary Table S* / Supplementary Fig. S*."""
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if re.match(r"^Supplementary Table S\d+\.", t):
            continue
        m = re.match(r"^Table S(\d+)\.\s*(.*)$", t)
        if m:
            p.text = f"Supplementary Table S{m.group(1)}. {m.group(2)}"
            continue
        m = re.match(r"^Figure S(\d+)\.\s*(.*)$", t, re.I)
        if m:
            p.text = f"Supplementary Fig. S{m.group(1)}. {m.group(2)}"


def _p_has_embedded_image(p_el) -> bool:
    return b'r:embed="' in etree.tostring(p_el)


def _embed_supplementary_figures(doc: Document) -> None:
    """Supplementary figures belong in Supplementary/Supplementary_Figures.docx, not the main manuscript."""
    print("Skipping supplementary figure embed in main manuscript (see Supplementary/ folder).")


def _audit(doc: Document) -> list[str]:
    issues: list[str] = []
    ref_i = next(i for i, p in enumerate(doc.paragraphs) if (p.text or "").strip() == "References")

    for i, p in enumerate(doc.paragraphs[:ref_i]):
        t = p.text or ""
        if re.search(r"(?<!Supplementary )Table S\d+", t):
            issues.append(f"para {i}: Table S* without 'Supplementary'")
        if re.search(r"(?<!Supplementary )Fig\. S\d+", t):
            issues.append(f"para {i}: Fig. S* without 'Supplementary'")
        if "Supplementary Fig.s" in t:
            issues.append(f"para {i}: invalid 'Supplementary Fig.s'")
        if re.search(r"Supplementary Fig\. S\d+[A-Z]", t):
            issues.append(f"para {i}: supplementary figure panel reference")

    in_refs = False
    n_refs = 0
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == "References":
            in_refs = True
            continue
        if in_refs:
            if t.startswith("Acknowledgements"):
                break
            if re.match(r"^\d+\.", t):
                n_refs += 1
                if not any(r.italic for r in p.runs if (r.text or "").strip()):
                    issues.append(f"ref {n_refs}: journal not italic")
                if " ." in t or "doi .org" in t:
                    issues.append(f"ref {n_refs}: stray spaces in journal/DOI")

    main_text = "\n".join(p.text or "" for p in doc.paragraphs[:ref_i])
    cited_tables: set[str] = set()
    cited_figs: set[str] = set()
    for m in re.finditer(r"Supplementary Tables? S(\d+)(?:[–\-—]S(\d+))?", main_text):
        cited_tables.add(m.group(1))
        if m.group(2):
            for k in range(int(m.group(1)), int(m.group(2)) + 1):
                cited_tables.add(str(k))
    for m in re.finditer(r"Supplementary Figs?\. S(\d+)(?:[–\-—]S(\d+))?", main_text):
        cited_figs.add(m.group(1))
        if m.group(2):
            for k in range(int(m.group(1)), int(m.group(2)) + 1):
                cited_figs.add(str(k))
    end_tables = set(re.findall(r"Supplementary Table S(\d+)\.", "\n".join(p.text or "" for p in doc.paragraphs)))
    end_figs = set(re.findall(r"Supplementary Fig\. S(\d+)\.", "\n".join(p.text or "" for p in doc.paragraphs)))

    for n in sorted(end_tables, key=int):
        if n not in cited_tables:
            issues.append(f"Supplementary Table S{n} listed but not cited in main text")
    for n in sorted(end_figs, key=int):
        if n not in cited_figs:
            issues.append(f"Supplementary Fig. S{n} listed but not cited in main text")

    return issues


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", type=Path, default=DEFAULT_DOCX)
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()
    path = args.docx.resolve()
    if not path.is_file():
        raise SystemExit(f"Not found: {path}")

    if not args.no_backup:
        bak = path.with_suffix(path.suffix + ".pre_srep_refs.bak")
        shutil.copy2(path, bak)
        print(f"Backup: {bak}")

    doc = Document(str(path))
    _rewrite_reference_list(doc)
    _replace_in_all(doc, SUPP_REPLACEMENTS)
    _fix_supplementary_paragraphs(doc)
    _remove_duplicate_legend_heading(doc)
    _normalize_supplementary_captions(doc)
    _embed_supplementary_figures(doc)
    doc.save(str(path))

    issues = _audit(Document(str(path)))
    if issues:
        print("Audit warnings:")
        for x in issues:
            print(f"  - {x}")
    else:
        print("Audit: references and supplementary citations OK for Scientific Reports.")
    print(f"Updated: {path}")


if __name__ == "__main__":
    main()
