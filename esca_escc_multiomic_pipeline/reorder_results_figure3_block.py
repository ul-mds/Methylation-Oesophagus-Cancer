"""
Place Figure 3 (caption + image) immediately after the Figure 2 image, then the
interpretive paragraph that begins 'Figure 3 compresses...', then Tables 2–5.
Reads more naturally: volcano → bar figure → interpretation → numeric tables.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from lxml import etree


def main():
    path = Path(__file__).resolve().parents[1] / "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"
    doc = Document(str(path))

    def find_prefix(prefix: str):
        for p in doc.paragraphs:
            if (p.text or "").strip().startswith(prefix):
                return p
        raise ValueError(f"Missing: {prefix!r}")

    p_fig2_cap = find_prefix("Figure 2.")
    p_fig3_compress = find_prefix("Figure 3 compresses these landscapes")
    p_fig3_cap = find_prefix("Figure 3. Top fast-death hypo-")

    # Walk from fig2 caption to locate fig2 image (first following para with embed)
    def next_sibling_p(el):
        n = el.getnext()
        while n is not None:
            if n.tag.endswith("}p"):
                return n
            n = n.getnext()
        return None

    el2 = p_fig2_cap._element
    el2_img = next_sibling_p(el2)
    if el2_img is None:
        raise RuntimeError("No paragraph after Figure 2 caption")

    el3_cap = p_fig3_cap._element
    el3_img = next_sibling_p(el3_cap)
    if el3_img is None or b"r:embed" not in etree.tostring(el3_img):
        raise RuntimeError("Expected Figure 3 caption to be followed by image paragraph")

    el_compress = p_fig3_compress._element

    parent = el2_img.getparent()
    if parent is None:
        raise RuntimeError("No body parent")

    # Detach Figure 3 caption + image from current location
    parent.remove(el3_cap)
    parent.remove(el3_img)

    # Insert after Figure 2 image
    el2_img.addnext(el3_cap)
    el3_cap.addnext(el3_img)

    # Ensure interpretive paragraph follows Figure 3 image (may already)
    n = el3_img.getnext()
    if n is not el_compress:
        parent.remove(el_compress)
        el3_img.addnext(el_compress)

    doc.save(str(path))
    print("Reordered Figure 3 block:", path)


if __name__ == "__main__":
    main()
