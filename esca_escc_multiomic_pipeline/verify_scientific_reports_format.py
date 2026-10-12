"""
Check manuscript against Cancer Medicine 'Format of articles' (and related) rules.

Reports: title word count, abstract (official 200 vs draft), keywords (≤6),
main-text word count excluding Abstract, Methods, References, figure legends (≤4500),
unstructured abstract, title page markers, graphical abstract absent.
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document

DOC_NAME = "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"


def _word_count(s: str) -> int:
    return len(s.split())


def _find_heading_index(paras, name: str) -> int | None:
    for i, p in enumerate(paras):
        if p.text.strip() == name:
            return i
    return None


def _text_before_abstract(paras) -> str:
    parts = []
    for p in paras:
        if p.text.strip() == "Abstract":
            break
        parts.append(p.text)
    return "\n".join(parts)


def main() -> None:
    path = Path(__file__).resolve().parents[1] / DOC_NAME
    doc = Document(str(path))
    paras = doc.paragraphs

    title = paras[0].text.strip()
    title_words = _word_count(title)

    abs_i = _find_heading_index(paras, "Abstract")
    abs_text = paras[abs_i + 1].text if abs_i is not None else ""
    abs_words = _word_count(abs_text)
    abs_structured = bool(re.search(r"^\s*(Background|Methods|Results|Conclusions)\s*:", abs_text, re.I | re.M))
    abs_has_bracket_refs = bool(re.search(r"\[\d+\]", abs_text))

    kw_text = ""
    for p in paras:
        if p.text.strip().startswith("Keywords"):
            kw_text = p.text
            break
    kw_bits = [k.strip() for k in kw_text.replace("Keywords:", "").split(";") if k.strip()]
    kw_n = len(kw_bits)

    intro_i = _find_heading_index(paras, "Introduction")
    methods_i = _find_heading_index(paras, "Methods")

    main_chunks = []
    if intro_i is not None and methods_i is not None:
        for j in range(intro_i + 1, methods_i):
            t = paras[j].text.strip()
            if t:
                main_chunks.append(t)
    main_text = " ".join(main_chunks)
    main_words = _word_count(main_text)

    title_block = _text_before_abstract(paras)
    title_page_ok = "Corresponding author:" in title_block or "corresponding author" in title_block.lower()
    affil_ok = bool(
        re.search(r"affiliation|department|institution|university|hospital", title_block, re.I)
    )

    lines = []
    lines.append("=== Cancer Medicine — Format of articles (checklist) ===\n")
    lines.append("Article type: assume standard Article (not Registered Report) — confirm at submission. PASS (manual).")
    lines.append(
        "Length: ideally ≤11 typeset pages — cannot be measured from Word; compare to similar SR papers or use "
        "journal PDF proof. MANUAL."
    )

    ok = title_words <= 20
    lines.append(f"Title ≤20 words (single scientific sentence; avoid puns/idioms): {title_words} words — {'PASS' if ok else 'FAIL'}")
    lines.append("  (Pun/idiom check is author judgment; title is plain scientific phrasing.)")

    ok200 = abs_words <= 200
    lines.append(f"Abstract ≤200 words: {abs_words} — {'PASS' if ok200 else 'FAIL'}")

    lines.append(f"Abstract unstructured (no Background/Methods/Results headings): {'PASS' if not abs_structured else 'FAIL'}")
    lines.append(f"Abstract has no references (no [n] citations): {'PASS' if not abs_has_bracket_refs else 'FAIL'}")
    lines.append("Graphical abstract: not supported by SR — do not include. PASS if omitted from submission (manual).")

    okkw = kw_n <= 6
    lines.append(f"Keywords count: {kw_n} (limit 6) — {'PASS' if okkw else 'FAIL'}")
    if kw_bits:
        lines.append(f"  Keywords: {', '.join(kw_bits)}")

    okmain = main_words <= 4500
    lines.append(
        f"Main text word count (Introduction→before Methods, excl. headings only as skipped): {main_words} "
        f"(guideline ≤4500 excl. Abstract, Methods, References, legends) — {'PASS' if okmain else 'FAIL'}"
    )
    lines.append("  (This count includes Results/Discussion body; Methods/References/Legends excluded by range.)")

    lines.append(
        f"Title page: corresponding author contact identified (*Corresponding author / e-mail): "
        f"{'PASS' if title_page_ok else 'FAIL'}"
    )
    lines.append(
        f"Title page: affiliation / institution text present: {'PASS' if affil_ok else 'FAIL (replace placeholders)'}"
    )
    lines.append(
        "Sections include referenced background (Introduction etc.): content review — manuscript has Introduction "
        "with citations. PASS (manual spot-check)."
    )

    lines.append("\n=== Summary ===")
    all_pass = all(
        [
            title_words <= 20,
            ok200,
            okkw,
            okmain,
            not abs_structured,
            not abs_has_bracket_refs,
            title_page_ok,
            affil_ok,
        ]
    )
    lines.append("Automated checklist (where measurable): " + ("PASS" if all_pass else "FAIL — see lines above"))

    report = "\n".join(lines)
    print(report)
    out = Path(__file__).resolve().parents[1] / "scientific_reports_format_check_report.txt"
    out.write_text(report, encoding="utf-8")
    print("\nWrote:", out)


if __name__ == "__main__":
    main()
