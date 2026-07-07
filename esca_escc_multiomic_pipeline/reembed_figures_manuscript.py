"""
Ensure each Figure N caption in the manuscript is immediately followed by its PNG
(embedded via python-docx). Fixes cases where captions exist but images were lost.

Also removes stray empty paragraphs between a figure caption and the next content
when no image was present (optional cleanup).

Run after: table renumbering / text edits that may have stripped drawings.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from lxml import etree


def _p_has_embedded_image(p_el) -> bool:
    return b'r:embed="' in etree.tostring(p_el)


def _insert_paragraph_after(paragraph: Paragraph) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)


def _ensure_image_after_caption(caption_para: Paragraph, png: Path, width: Inches) -> None:
    """Place PNG immediately after the caption paragraph, replacing any existing embedded image."""
    cap_el = caption_para._p
    nxt = cap_el.getnext()
    if nxt is not None and nxt.tag.endswith("}p") and _p_has_embedded_image(nxt):
        next_para = Paragraph(nxt, caption_para._parent)
        for r in list(next_para.runs):
            r._r.getparent().remove(r._r)
        next_para.add_run().add_picture(str(png), width=width)
        return

    if nxt is not None and nxt.tag.endswith("}p"):
        next_para = Paragraph(nxt, caption_para._parent)
        txt = (next_para.text or "").strip()
        if not txt:
            for r in list(next_para.runs):
                r._r.getparent().remove(r._r)
            next_para.add_run().add_picture(str(png), width=width)
            return

    pic_p = _insert_paragraph_after(caption_para)
    pic_p.add_run().add_picture(str(png), width=width)


def _find_caption_paragraph(doc: Document, prefix: str) -> Paragraph:
    for p in doc.paragraphs:
        if (p.text or "").strip().startswith(prefix):
            return p
    raise ValueError(f"Missing caption paragraph starting with: {prefix!r}")


def main():
    study = Path(__file__).resolve().parents[1]
    docx_path = study / "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"
    plot_dir = study / "Plots"

    specs: list[tuple[str, str, float]] = [
        ("Figure 1.", "01_cross_apply_cindex.png", 6.2),
        ("Figure 2.", "02_fastdeath_methylation_volcano_panel.png", 6.6),
        ("Figure 3.", "03_fastdeath_top_hypo_hyper_panel.png", 6.6),
        ("Figure 4.", "04_KM_methylation_60mo_panel.png", 6.6),
        ("Figure 5.", "05_mutation_module_heatmap.png", 6.0),
        ("Figure 6.", "06_interaction_screen.png", 6.0),
    ]

    for _, stem, _ in specs:
        p = plot_dir / stem
        if not p.is_file():
            raise FileNotFoundError(f"Missing plot: {p}")

    doc = Document(str(docx_path))
    for prefix, stem, w in specs:
        cap = _find_caption_paragraph(doc, prefix)
        _ensure_image_after_caption(cap, plot_dir / stem, Inches(w))

    doc.save(str(docx_path))

    try:
        from .patch_manuscript_figures import _prune_orphan_image_parts

        _prune_orphan_image_parts(docx_path)
    except Exception:
        pass

    print("Re-embedded figures:", docx_path)


if __name__ == "__main__":
    main()
