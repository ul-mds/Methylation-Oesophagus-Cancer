"""
Place each figure immediately after the Results paragraph that first establishes
its interpretive context. Remove the trailing consolidated 'Figures' gallery.

Order:
  • After transfer paragraph → Figure 1 → (existing Table 1 caption follows)
  • After volcano paragraph (Figure 2 text) → Figure 2 → Figure 3 paragraph → Figure 3 → Tables 2–5
  • After Table 6 caption → Figure 4 → somatic paragraphs → Figure 5 → Table 7 → Figure 6
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.text.paragraph import Paragraph


def _insert_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def _insert_figure_block_after(
    paragraph: Paragraph, caption: str, png: Path, width: Inches | float
) -> Paragraph:
    """Insert caption paragraph then image paragraph; return the image paragraph."""
    cap = _insert_paragraph_after(paragraph, caption)
    pic_p = _insert_paragraph_after(cap, "")
    pic_p.add_run().add_picture(str(png), width=width)
    return pic_p


def _find_paragraph(doc: Document, prefix: str) -> Paragraph:
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t.startswith(prefix):
            return p
    raise ValueError(f"No paragraph starting with: {prefix!r}")


def _figures_already_inlined(doc: Document) -> bool:
    """True if Figure 1 caption sits between the transfer paragraph and Table 1."""
    p26_i = None
    t1_i = None
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t.startswith("We trained penalized Cox models"):
            p26_i = i
        if p26_i is not None and t.startswith("Table 1.") and "Concordance" in t:
            t1_i = i
            break
    if p26_i is None or t1_i is None:
        return False
    for j in range(p26_i + 1, t1_i):
        if (doc.paragraphs[j].text or "").strip().startswith("Figure 1."):
            return True
    return False


def _remove_figures_gallery(doc: Document) -> None:
    """Delete from 'Figures' heading through paragraph immediately before 'Discussion'."""
    start = None
    disc = None
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t.startswith("Figures") and "panels below" in t:
            start = i
        if start is not None and t == "Discussion":
            disc = i
            break
    if start is None or disc is None:
        return
    for j in range(disc - 1, start - 1, -1):
        el = doc.paragraphs[j]._element
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)


def main():
    study = Path(__file__).resolve().parents[1]
    docx_path = study / "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"
    plot_dir = study / "Plots"

    for stem in (
        "01_cross_apply_cindex.png",
        "02_fastdeath_methylation_volcano_panel.png",
        "03_fastdeath_top_hypo_hyper_panel.png",
        "04_KM_methylation_60mo_panel.png",
        "05_mutation_module_heatmap.png",
        "06_interaction_screen.png",
    ):
        if not (plot_dir / stem).is_file():
            raise FileNotFoundError(f"Missing plot: {plot_dir / stem}")

    captions = {
        1: (
            "Figure 1. Cross-apply performance (C-index) using 60-month landmark OS. "
            "Histology-specific models discriminate within histology but do not transfer across histologies."
        ),
        2: (
            "Figure 2. Fast-death differential methylation by histology (volcano). Left: ESCA/EAC; right: ESCC. "
            "Each point is a locus with Δβ = mean(β_fast)−mean(β_other) and y-axis = −log10(q_diff)."
        ),
        3: (
            "Figure 3. Top fast-death hypo- and hypermethylated loci (Δβ) in ESCA/EAC (left) and ESCC (right)."
        ),
        4: (
            "Figure 4. Kaplan–Meier curves (60-month landmark OS) stratified by median methylation. "
            "Upper row: ESCA/EAC, three probes with strongest methylation–OS signal (univariate Cox on OS60). "
            "Lower row: ESCC, three probes with strongest fast-death differential methylation (q_diff)."
        ),
        5: (
            "Figure 5. Exploratory mechanistic follow-up: mutation-associated shifts in expression-derived "
            "module scores (Δ=mut−WT) stratified by histology."
        ),
        6: (
            "Figure 6. Mutation×histology interaction screen on 60-month landmark OS (nominal p-values)."
        ),
    }

    doc = Document(str(docx_path))

    if _figures_already_inlined(doc):
        print("Figures already inlined; skipping:", docx_path)
        return

    # Remove old end-of-manuscript figure block (captions + images).
    _remove_figures_gallery(doc)

    # Re-find anchors after removal (indices shifted).
    p26 = _find_paragraph(doc, "We trained penalized Cox models")
    p29 = _find_paragraph(doc, "Genome-wide contrasts revealed divergent")
    p30 = _find_paragraph(doc, "Figure 3 compresses these landscapes")
    p41 = _find_paragraph(doc, "Table 6. EAC/ESCA: univariate Cox")
    p44 = _find_paragraph(doc, "Because mutation×histology interaction tests")

    # Insert from bottom to top so we do not rely on stale indices.
    tail = _insert_figure_block_after(
        p44, captions[6], plot_dir / "06_interaction_screen.png", Inches(6.0)
    )
    _insert_figure_block_after(
        tail, captions[5], plot_dir / "05_mutation_module_heatmap.png", Inches(6.0)
    )

    _insert_figure_block_after(
        p41, captions[4], plot_dir / "04_KM_methylation_60mo_panel.png", Inches(6.6)
    )

    _insert_figure_block_after(
        p30, captions[3], plot_dir / "03_fastdeath_top_hypo_hyper_panel.png", Inches(6.6)
    )

    _insert_figure_block_after(
        p29, captions[2], plot_dir / "02_fastdeath_methylation_volcano_panel.png", Inches(6.6)
    )

    _insert_figure_block_after(
        p26, captions[1], plot_dir / "01_cross_apply_cindex.png", Inches(6.2)
    )

    doc.save(str(docx_path))

    # Drop unused image relationships / media from prior figure gallery.
    try:
        from .patch_manuscript_figures import _prune_orphan_image_parts

        _prune_orphan_image_parts(docx_path)
    except Exception:
        pass

    print("Inlined figures into Results:", docx_path)


if __name__ == "__main__":
    main()
