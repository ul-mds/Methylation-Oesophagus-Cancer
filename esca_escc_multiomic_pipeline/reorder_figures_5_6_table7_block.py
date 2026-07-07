"""
Place Figure 5 (caption + image), then Table 7 caption, then Figure 6 (caption + image),
after the bridge paragraph that introduces the somatic/mutation subsection.

Matches citation order in the opening arc (Figures 5–6; Table 7) and keeps Table 7 next
to the Figure 5 heatmap discussion.
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
        raise ValueError(f"Missing paragraph: {prefix!r}")

    def next_sibling_p(el):
        n = el.getnext()
        while n is not None:
            if n.tag.endswith("}p"):
                return n
            n = n.getnext()
        return None

    p_bridge = find_prefix("Because mutation×histology interaction tests are event-limited")
    p_fig5 = find_prefix("Figure 5. Exploratory mechanistic follow-up:")
    p_fig6 = find_prefix("Figure 6. Mutation×histology interaction screen")
    p_tbl7 = find_prefix("Table 7. Exploratory mechanistic follow-up:")

    el5c, el6c, el7c = p_fig5._element, p_fig6._element, p_tbl7._element
    el5i = next_sibling_p(el5c)
    el6i = next_sibling_p(el6c)
    if el5i is None or el6i is None:
        raise RuntimeError("Expected image paragraphs after Figure 5 and 6 captions")
    if b"r:embed" not in etree.tostring(el5i) or b"r:embed" not in etree.tostring(el6i):
        raise RuntimeError("Expected r:embed in paragraphs after Figure 5 and 6 captions")

    parent = el5c.getparent()
    if parent is None or parent is not el6c.getparent() or parent is not el7c.getparent():
        raise RuntimeError("Figure 5/6/Table 7 blocks are not siblings in body")

    # Detach in document order (5, 5img, 6, 6img, 7) or any order — remove all first
    for el in (el5c, el5i, el6c, el6i, el7c):
        parent.remove(el)

    el_br = p_bridge._element
    # Insert: Fig5 cap, Fig5 img, Table 7, Fig6 cap, Fig6 img
    el_br.addnext(el5c)
    el5c.addnext(el5i)
    el5i.addnext(el7c)
    el7c.addnext(el6c)
    el6c.addnext(el6i)

    p_bridge.text = (
        "Because mutation×histology interaction tests are event-limited in moderate samples, we paired "
        "histology-stratified transcriptional context with a nominal interaction discovery map. Figure 5 "
        "summarizes differences in expression-derived module scores between mutant and wild-type tumors "
        "(Δ=mutant−wild-type), including histology-partitioned movement on EMT and immune/interferon axes "
        "for selected alterations. Table 7 lists the strongest mutation-to-module or mutation-to-expression "
        "associations underlying that heatmap. Figure 6 then presents the interaction landscape (Cox β for "
        "the product term versus −log10 p): labeled genes highlight where histology appears to modify the "
        "mutation–survival slope and warrant independent replication under larger event counts. Read "
        "sequentially, Figure 5 and Table 7 offer testable transcriptional hypotheses integrable with the "
        "methylation programs above, while Figure 6 prioritizes genomic candidates for follow-up."
    )

    doc.save(str(path))
    print("Reordered Figure 5 → Table 7 → Figure 6 block:", path)


if __name__ == "__main__":
    main()
