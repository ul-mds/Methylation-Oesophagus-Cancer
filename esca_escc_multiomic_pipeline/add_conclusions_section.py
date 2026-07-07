"""
Insert a Heading 1 'Conclusions' section with two substantive paragraphs immediately before References.

Idempotent: skips if a Heading 1 paragraph titled 'Conclusions' already precedes References.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

DOC_NAME = "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"

BODY_1 = (
    "Taken together, this work supports a simple but consequential principle for esophageal carcinoma omics: "
    "EAC and ESCC should not be treated as interchangeable prognostic substrates. Histology-specific models retain "
    "discrimination within lineage yet fail when coefficients trained in one histology are applied to the other, "
    "which cautions against silent pooling in signature development and in secondary use of public resources. "
    "The fast-death–anchored methylation contrasts reinforce that aggressive clinical trajectories map to distinct "
    "epigenetic programmes by histology rather than to a single pan-oesophageal programme, and that probe-level "
    "survival separation can still be demonstrated when hypotheses are framed in lineage-resolved terms."
)

BODY_2 = (
    "For translation, the priority is prospective validation in external cohorts with matched assays and explicit "
    "histology annotation, including independent HM450 or successor platforms and clinically annotated survival. "
    "Exploratory somatic interaction and module summaries should be read as hypothesis generators that require larger "
    "event counts and orthogonal functional follow-up. Reporting esophageal omics findings with histology as a first-class "
    "stratum—rather than as a post hoc label—would align statistical practice with the molecular architecture already "
    "documented for these diseases and would make multi-centre comparisons and meta-analyses more interpretable."
)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    path = root / DOC_NAME
    doc = Document(str(path))

    for p in doc.paragraphs:
        if (p.text or "").strip() == "Conclusions" and p.style.name.startswith("Heading 1"):
            print("Conclusions already present; skipping.")
            return

    ref_para = None
    for p in doc.paragraphs:
        if (p.text or "").strip() == "References" and p.style.name.startswith("Heading 1"):
            ref_para = p
            break
    if ref_para is None:
        raise SystemExit("Could not find References heading.")

    ref_el = ref_para._p
    h_el = OxmlElement("w:p")
    ref_el.addprevious(h_el)
    h_p = Paragraph(h_el, doc._body)
    h_p.style = doc.styles["Heading 1"]
    h_p.add_run("Conclusions")

    p1_el = OxmlElement("w:p")
    h_el.addnext(p1_el)
    p1 = Paragraph(p1_el, doc._body)
    p1.style = doc.styles["Normal"]
    p1.add_run(BODY_1)

    p2_el = OxmlElement("w:p")
    p1_el.addnext(p2_el)
    p2 = Paragraph(p2_el, doc._body)
    p2.style = doc.styles["Normal"]
    p2.add_run(BODY_2)

    doc.save(str(path))
    print("Inserted Conclusions before References:", path)


if __name__ == "__main__":
    main()
