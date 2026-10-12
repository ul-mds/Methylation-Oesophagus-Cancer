"""
Reformat the manuscript toward Cancer Medicine submission guidelines:
https://onlinelibrary.wiley.com/page/journal/20457634/homepage/forauthors.html

- Unstructured abstract (≤200 words, no references)
- Main text: Introduction → Results → Discussion (includes former Conclusion) → Methods
- References in Nature-style numbered list; mandatory end sections
- Figure legends block; note on ≤8 display items

Idempotent if a paragraph starts with 'Data availability'.
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph


def _ptext(el) -> str:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    return "".join(t.text or "" for t in el.findall(".//w:t", ns)).strip()


def _find_body_index(body, predicate) -> int:
    for i, el in enumerate(body):
        if el.tag.endswith("}p") and predicate(_ptext(el)):
            return i
    raise ValueError("paragraph not found")


def _already_srep_formatted(body) -> bool:
    for el in body:
        if el.tag.endswith("}p") and _ptext(el).startswith("Data availability"):
            return True
    return False


def main():
    path = Path(__file__).resolve().parents[1] / "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"
    doc = Document(str(path))
    body = doc.element.body

    if _already_srep_formatted(body):
        print("Already formatted (Data availability found):", path)
        return

    # --- Collect figure legend strings before any edits ---
    legend_lines = []
    for el in body:
        if el.tag.endswith("}p"):
            t = _ptext(el)
            if re.match(r"^Figure \d+\.", t):
                legend_lines.append(t)

    # --- 1. Unstructured abstract ---
    abs_start = _find_body_index(body, lambda t: t == "Abstract")
    p_abs_content = body[abs_start + 1]
    if not _ptext(p_abs_content).startswith("Background:"):
        raise RuntimeError("Expected structured abstract (Background: …).")

    unstructured = (
        "Esophageal adenocarcinoma (EAC) and esophageal squamous cell carcinoma (ESCC) are often pooled in "
        "secondary analyses despite distinct biology. Using the TCGA esophageal carcinoma (ESCA, Firehose Legacy) "
        "cohort, we asked whether multi-omic prognostic structure transfers across histologies when overall survival "
        "is analysed on a 60-month landmark with HM450 methylation and expression features. Histology-specific Cox "
        "models discriminated strongly within EAC and within ESCC yet collapsed when cross-applied, whereas a pooled "
        "model with histology-stratified baseline hazards achieved C-index 0.803—indicating that prognosis is "
        "recoverable in a combined cohort only when histology is encoded in the hazard. Among decedents, a "
        "fast-death tail within each histology uncovered divergent methylation programs; representative loci included "
        "SIGLEC14, EPHB1, ZNF214, and RAB11FIP3 in EAC and SHISA3, C7orf13, PRKG1, and SEMA3E in ESCC, with "
        "methylation linked to interpretable 60-month survival separation. Exploratory mutation-by-histology "
        "interactions and expression-module shifts contextualise somatic variation. These findings support treating "
        "EAC and ESCC as separate prognostic entities for omics-derived risk and motivate histology-matched "
        "external validation."
    )
    if len(unstructured.split()) > 200:
        raise RuntimeError("Abstract exceeds 200 words.")

    for p in doc.paragraphs:
        if p._element is p_abs_content:
            p.text = unstructured
            break

    to_remove = []
    j = abs_start + 2
    for _ in range(3):
        if j < len(body) and body[j].tag.endswith("}p"):
            to_remove.append(body[j])
            j += 1
    for el in to_remove:
        el.getparent().remove(el)

    body = doc.element.body
    sect_pr = body[-1]
    if not sect_pr.tag.endswith("}sectPr"):
        raise RuntimeError("Expected sectPr last.")

    children = [c for c in body if c is not sect_pr]

    def idx(pred):
        for i, el in enumerate(children):
            if el.tag.endswith("}p") and pred(_ptext(el)):
                return i
        raise ValueError("index")

    i_kw = idx(lambda t: t.startswith("Keywords:"))
    i_intro = idx(lambda t: t == "Introduction")
    i_meth = idx(lambda t: t == "Materials and Methods")
    i_res = idx(lambda t: t == "Results")
    i_disc = idx(lambda t: t == "Discussion")
    i_conc = idx(lambda t: t == "Conclusion")
    i_refs = idx(lambda t: t == "References")
    i_supp = idx(lambda t: t == "Supplementary Materials")

    front = children[0 : i_kw + 1]
    intro = children[i_intro:i_meth]
    methods = children[i_meth:i_res]
    results = children[i_res:i_disc]
    discussion = children[i_disc:i_conc]
    conclusion_paras = children[i_conc + 1 : i_refs]
    refs_and_supp = children[i_refs:]

    conc_heading = children[i_conc]
    if _ptext(conc_heading) != "Conclusion":
        raise RuntimeError("Expected Conclusion heading before References.")

    new_order = front + intro + results + discussion + conclusion_paras + methods + refs_and_supp

    for c in list(body):
        if c is not sect_pr:
            body.remove(c)
    for c in new_order:
        body.append(c)
    body.append(sect_pr)

    body = doc.element.body
    for el in body:
        if el.tag.endswith("}p") and _ptext(el) == "Materials and Methods":
            for p in doc.paragraphs:
                if p._element is el:
                    p.text = "Methods"
                    break
            break

    # --- Nature-style references + end matter (after References heading, before Supplementary) ---
    ref_lines = [
        "Cancer Genome Atlas Research Network. Integrated genomic characterization of oesophageal carcinoma. "
        "Nature 541, 169-175 (2017).",
        "Li, B. et al. Identification of DNA methylation-driven genes in esophageal squamous cell carcinoma: "
        "a study based on The Cancer Genome Atlas. Cancer Cell Int. 19, 82 (2019).",
        "Chen, L. et al. Identification of prognostic risk model based on DNA methylation-driven genes in "
        "esophageal adenocarcinoma. Biomed. Res. Int. 2021, 6690910 (2021).",
        "Cao, W. et al. Multi-faceted epigenetic dysregulation of gene expression promotes esophageal squamous "
        "cell carcinoma. Nat. Commun. 11, 3675 (2020).",
        "Xi, Y. et al. Multi-omic characterization of genome-wide abnormal DNA methylation reveals diagnostic "
        "and prognostic markers for esophageal squamous-cell carcinoma. Sig. Transduct. Target. Ther. 7, 53 "
        "(2022).",
        "Lyu, Q. et al. Integrated bioinformatics analysis of differences between EAC and ESCC. BMC Cancer "
        "25, 1668 (2025).",
        "Pinto, R. et al. Molecular stratification of esophageal adenocarcinoma: implications for prognosis and "
        "treatment strategy. Oncogene (2025); https://doi.org/10.1038/s41388-025-03650-3",
        "Tang, J. et al. DNA methylation in esophageal cancer: technological advances and early detection "
        "clinical applications. Front. Oncol. 15, 1543190 (2025).",
        "Wang, G. et al. Analysis of the associations between DNA methylation and clinical features reveals the "
        "potential oncogenic effect of CFAP52 in esophageal squamous cell cancer. Cancer Cell Int. 25, 394 "
        "(2025).",
        "Cui, H. et al. The consensus molecular subtypes of esophageal squamous cell carcinoma. "
        "Sig. Transduct. Target. Ther. 11, 65 (2026).",
        "Zheng, Y. et al. Genome-wide DNA methylation and gene expression profiling characterizes molecular "
        "subtypes of esophagus squamous cell carcinoma for predicting patient survival and immunotherapy efficacy. "
        "Cancers 14, 4970 (2022).",
        "Zhang, X. et al. Multi-omics analysis unveils a four-gene prognostic signature in esophageal squamous "
        "carcinoma and the therapeutic potential of PKP1. BMC Cancer 25, 777 (2025).",
        "Oesophageal Cancer Clinical and Molecular Stratification (OCCAMS) Consortium et al. Disentangling "
        "oncogenic amplicons in esophageal adenocarcinoma. Nat. Commun. 15, 4074 (2024).",
        "Chang, Z. et al. Construction of a multimodal esophageal cancer prognostic model with screening of key "
        "genes and investigation of their biological functions. Discov. Oncol. 16, 1560 (2025).",
        "Tang, P. et al. Integrated machine learning developed a prognosis-related gene signature to predict "
        "prognosis in oesophageal squamous cell carcinoma. J. Cell. Mol. Med. 28, e70171 (2024).",
        "Tuersun, A. et al. Interpretable machine learning models to predict survival in esophageal cancer: "
        "a study based on the SEER database and external validation in China. Front. Physiol. 16, 1665383 "
        "(2025).",
        "Pu, Y. et al. Estimating the prognosis of esophageal squamous cell carcinoma based on The Cancer Genome "
        "Atlas (TCGA) of m6A methylation-associated genes. J. Gastrointest. Oncol. 13, 1-12 (2022).",
        "Zhao, D. et al. Advances in multi-omics for esophageal squamous cell carcinoma: diagnostic, prognostic, "
        "and therapeutic perspectives. Protein Cell (2026); https://doi.org/10.1093/procel/pwag005",
        "Kunene, V. et al. Prognostic markers in oesophageal and gastric cancer review. Are they ready for clinical "
        "practice? ESMO Gastrointest. Oncol. 1, 100091 (2024).",
        "Xu, Y. et al. DNA methylation markers in esophageal cancer. Front. Genet. 15, 1354195 (2024).",
    ]

    body = doc.element.body
    i_refs = _find_body_index(body, lambda t: t == "References")
    i_supp = _find_body_index(body, lambda t: t == "Supplementary Materials")

    kill = []
    for k in range(i_refs + 1, i_supp):
        kill.append(body[k])
    for el in kill:
        el.getparent().remove(el)

    body = doc.element.body
    i_refs = _find_body_index(body, lambda t: t == "References")
    i_supp = _find_body_index(body, lambda t: t == "Supplementary Materials")
    ref_head = body[i_refs]
    prev = ref_head
    parent = doc._body
    for n, line in enumerate(ref_lines, start=1):
        new_p = OxmlElement("w:p")
        prev.addnext(new_p)
        Paragraph(new_p, parent).add_run(f"{n}. {line}")
        prev = new_p

    body = doc.element.body
    i_supp = _find_body_index(body, lambda t: t == "Supplementary Materials")
    anchor = body[i_supp].getprevious()
    if anchor is None:
        raise RuntimeError("No element before Supplementary.")

    sections = [
        ("Acknowledgements", "The authors thank the TCGA Research Network and the cBioPortal team for curated "
         "public data. Grant numbers and personal acknowledgements may be inserted here."),
        ("Author contributions", "Conceptualisation, methodology, software, formal analysis, investigation, "
         "data curation, writing—original draft, writing—review and editing: [assign CRediT roles]. All authors "
         "approved the final version."),
        ("Data availability", "TCGA esophageal carcinoma (ESCA) Firehose Legacy data were accessed through "
         "cBioPortal (https://www.cbioportal.org/study/summary?id=esca_tcga). Analysis outputs supporting figures "
         "and tables are in the accompanying project directories (result/ and Plots/) and should be assigned a "
         "public repository DOI on acceptance."),
        ("Competing interests", "The authors declare no competing interests."),
        ("Additional information (display items)", "Cancer Medicine expects at most eight combined figures and "
         "tables in the main article. This draft lists six figures and seven tables; before submission, move excess "
         "items to Supplementary Information (or merge panels) and cite as Supplementary Fig./Table."),
    ]

    prev = anchor
    for title, blurb in sections:
        t_el = OxmlElement("w:p")
        prev.addnext(t_el)
        pt = Paragraph(t_el, parent)
        pt.style = "Heading 1"
        pt.add_run(title)
        prev = t_el
        b_el = OxmlElement("w:p")
        prev.addnext(b_el)
        Paragraph(b_el, parent).add_run(blurb)
        prev = b_el

    leg_h = OxmlElement("w:p")
    prev.addnext(leg_h)
    _lp = Paragraph(leg_h, parent)
    _lp.style = "Heading 1"
    _lp.add_run("Figure legends")
    prev = leg_h
    for leg in legend_lines:
        lp = OxmlElement("w:p")
        prev.addnext(lp)
        Paragraph(lp, parent).add_run(leg)
        prev = lp

    doc.save(str(path))
    print("Cancer Medicine–oriented formatting complete:", path)


if __name__ == "__main__":
    main()
