"""
Reformat Manuscript First.docx to BMC Cancer Research article section order only.
https://link.springer.com/journal/12885/submission-guidelines/research-article

Format-only: headings, section order, structured abstract split of existing sentences,
Declarations subheadings, abbreviation list, Vancouver reference punctuation.
No scientific rewriting. All insertions/renames are highlighted yellow.
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
YELLOW = WD_COLOR_INDEX.YELLOW

from esca_escc_multiomic_pipeline.study_paths import study_root

ROOT = study_root()
SRC = ROOT / "Manuscript First.docx"


def ptext(el) -> str:
    return "".join(t.text or "" for t in el.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")).strip()


def find_para(doc: Document, exact: str) -> Paragraph:
    for p in doc.paragraphs:
        if p.text.strip() == exact:
            return p
    raise ValueError(f"Paragraph not found: {exact!r}")


def find_para_startswith(doc: Document, prefix: str) -> Paragraph:
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefix):
            return p
    raise ValueError(f"Paragraph starting with {prefix!r} not found")


def clear_runs(paragraph: Paragraph) -> None:
    p = paragraph._p
    for child in list(p):
        if child.tag == qn("w:r") or child.tag == qn("w:hyperlink"):
            p.remove(child)


def set_highlighted(paragraph: Paragraph, text: str) -> None:
    clear_runs(paragraph)
    run = paragraph.add_run(text)
    run.font.highlight_color = YELLOW


def insert_after(paragraph: Paragraph, text: str, style: str | None = None) -> Paragraph:
    new_el = paragraph._p.makeelement(qn("w:p"), {})
    paragraph._p.addnext(new_el)
    new_p = Paragraph(new_el, paragraph._parent)
    if style:
        new_p.style = style
    run = new_p.add_run(text)
    run.font.highlight_color = YELLOW
    return new_p


def insert_before(paragraph: Paragraph, text: str, style: str | None = None) -> Paragraph:
    new_el = paragraph._p.makeelement(qn("w:p"), {})
    paragraph._p.addprevious(new_el)
    new_p = Paragraph(new_el, paragraph._parent)
    if style:
        new_p.style = style
    run = new_p.add_run(text)
    run.font.highlight_color = YELLOW
    return new_p


def move_block_before(doc: Document, start_text: str, end_text: str, before_text: str) -> None:
    start = find_para(doc, start_text)._element
    end = find_para(doc, end_text)._element
    target = find_para(doc, before_text)._element
    block = []
    el = start
    while el is not None and el is not end:
        nxt = el.getnext()
        block.append(el)
        el = nxt
    for el in block:
        target.addprevious(el)


def heading1_elements(doc: Document) -> set:
    return {p._element for p in doc.paragraphs if p.style.name == "Heading 1"}


def collect_until_next_h1(start_el, h1_els: set):
    """Heading 1 element plus following body until next Heading 1."""
    block = [start_el]
    el = start_el.getnext()
    while el is not None:
        if el in h1_els and el is not start_el:
            break
        nxt = el.getnext()
        block.append(el)
        el = nxt
    return block


def vancouverize_original(original: str) -> str:
    m = re.match(r"^(\d+)\.\s+(.*)$", original.strip(), re.S)
    if not m:
        return original
    num, rest = m.group(1), m.group(2).strip()

    doi = ""
    dm = re.search(r";\s*(https?://doi\.org/\S+)", rest)
    if dm:
        doi = dm.group(1).rstrip(".")
        rest = rest[: dm.start()].strip()
    rest = rest.rstrip(".")

    year = ""
    ym = re.search(r"\((\d{4})\)\s*$", rest)
    if ym:
        year = ym.group(1)
        rest = rest[: ym.start()].rstrip(" ,;")

    # Split authors vs title: authors end at ". " followed by title starting with capital
    # Pattern: "Last, I. I., Last, I. & Last, I. Title. Journal vol, pages"
    # or "Network. Title."
    # Find year-independent: last ". " before journal volume pattern `Journal. VOL, PAGES`
    vol_m = re.search(r"\s(\d[\d]*)\s*,\s*([\d]+(?:-\d+)?)\s*$", rest)
    journal_article = None
    pages = None
    volume = None
    if vol_m:
        volume = vol_m.group(1)
        pages = vol_m.group(2).replace("-", "–")
        head = rest[: vol_m.start()].rstrip()
        # head = authors. Title. Journal
        # split at last " ." that is journal
        parts = head.rsplit(". ", 1)
        if len(parts) == 2:
            authors_title, journal = parts
            # authors vs title: first ". " after author names is fragile for "et al."
            # Use: authors contain "et al." or initials pattern, then title.
            am = re.match(
                r"^(?P<authors>.+?)\.\s+(?P<title>.+)$",
                authors_title,
            )
            if am:
                authors = compress_authors(am.group("authors"))
                title = am.group("title").rstrip(".")
                journal = journal.rstrip(".")
                out = f"{num}. {authors}. {title}. {journal}."
                if year:
                    out += f" {year}"
                if volume:
                    out += f";{volume}"
                    if pages:
                        out += f":{pages}"
                out += "."
                if doi:
                    out += f" {doi}."
                return out
    # no volume/pages (e.g. Protein Cell (2026))
    am = re.match(r"^(?P<authors>.+?)\.\s+(?P<title>.+)$", rest)
    if am:
        authors = compress_authors(am.group("authors"))
        title_journal = am.group("title").rstrip(".")
        # last ". " separates title from journal if present
        if ". " in title_journal:
            title, journal = title_journal.rsplit(". ", 1)
            out = f"{num}. {authors}. {title}. {journal}."
        else:
            out = f"{num}. {authors}. {title_journal}."
        if year:
            out = out.rstrip(".") + f" {year}."
        if doi:
            out = out.rstrip(".") + f" {doi}."
        return out
    return original


def compress_authors(authors: str) -> str:
    authors = authors.replace(" & ", ", ")
    authors = re.sub(r"\b([A-Z])\.(?=\s|$)", r"\1", authors)
    authors = re.sub(r"\s+,", ",", authors)
    authors = re.sub(r",\s+", ", ", authors)
    authors = re.sub(r"\s{2,}", " ", authors).strip()
    # "et al" without period from stripping: restore
    authors = re.sub(r"\bet al$", "et al.", authors)
    authors = re.sub(r"\bet al\.(?=\.)", "et al", authors)
    return authors


def main() -> None:
    if not SRC.exists():
        raise FileNotFoundError(SRC)
    bak = SRC.with_suffix(".docx.pre_bmc_format.bak")
    shutil.copy2(SRC, bak)
    doc = Document(str(SRC))

    # 1) Methods before Results (BMC order: Background, Methods, Results, Discussion, Conclusions)
    move_block_before(doc, "Methods", "References", "Results")

    # 3) Introduction -> Background
    intro = find_para(doc, "Introduction")
    intro.style = "Heading 1"
    set_highlighted(intro, "Background")

    # 4) Structured abstract from existing sentences only
    abs_body = None
    for p in doc.paragraphs:
        if p.text.startswith("Esophageal adenocarcinoma (EAC) and esophageal squamous cell carcinoma (ESCC) are biologically distinct"):
            abs_body = p
            break
    if abs_body is None:
        raise RuntimeError("Abstract body not found")
    original_abs = abs_body.text.strip()
    # sentence split keeping original wording
    s1 = "Esophageal adenocarcinoma (EAC) and esophageal squamous cell carcinoma (ESCC) are biologically distinct, yet many prognostic omics studies analyse them as a single cohort."
    s2 = "Using TCGA ESCA (Firehose Legacy) data under a 60-month truncated overall survival horizon (OS60), we tested whether multi-omic risk scores trained in one histology retain prognostic discrimination in the other."
    s3 = "Within-lineage penalized Cox models showed high apparent (in-sample) discrimination (C-index 0.894 for EAC and 0.901 for ESCC), but cross-application of fixed coefficients collapsed to near-random performance (0.469 and 0.450)."
    s4 = "A pooled model with histology-stratified baseline hazards recovered cohort-wide separation (C-index 0.803), underscoring that meaningful prognostic structure exists but is not interchangeable across lineages."
    s5 = "We therefore reframed discovery around outcome-anchored methylation contrasts, defining a short-survival (fast-death) tail among decedents and comparing these cases with all other samples of the same histology."
    s6 = "This identified separate candidate epigenetic patterns in EAC and ESCC rather than a shared pan-esophageal signature."
    s7 = "Exploratory mutation-by-histology screens suggested context-dependent somatic associations."
    s8 = "Modular checks in external GEO cohorts supported biological concordance for selected expression modules but did not reproduce TCGA fast-death methylation signals on overall survival."
    s9 = "These findings argue for histology-aware reporting and assay-matched external validation before omics-based esophageal risk stratification is treated as clinically actionable."
    reconstructed = " ".join([s1, s2, s3, s4, s5, s6, s7, s8, s9])
    if reconstructed != original_abs:
        raise RuntimeError("Abstract sentence split does not match original abstract; aborting to avoid content change.")

    set_highlighted(abs_body, "Background: " + s1)
    p = insert_after(abs_body, "Methods: " + s2 + " " + s5, "Normal")
    p = insert_after(p, "Results: " + s3 + " " + s4 + " " + s6 + " " + s7 + " " + s8, "Normal")
    insert_after(p, "Conclusions: " + s9, "Normal")

    # 5) Keywords as its own heading
    kw = None
    for para in doc.paragraphs:
        if para.text.startswith("Keywords:"):
            kw = para
            break
    if kw is None:
        raise RuntimeError("Keywords not found")
    kw_text = kw.text.replace("Keywords:", "", 1).strip()
    kw.style = "Heading 1"
    set_highlighted(kw, "Keywords")
    insert_after(kw, kw_text, "Normal")

    # 6) Conclusions heading before existing closing paragraph of Discussion
    conc = find_para_startswith(doc, "Prognostic models intended for use across EAC and ESCC")
    insert_before(conc, "Conclusions", "Heading 1")

    # 7) List of abbreviations (terms already used in the manuscript)
    ack = find_para(doc, "Acknowledgements")
    insert_before(ack, "List of abbreviations", "Heading 1")
    abbr = (
        "EAC: esophageal adenocarcinoma; ESCC: esophageal squamous cell carcinoma; "
        "TCGA: The Cancer Genome Atlas; ESCA: esophageal carcinoma (TCGA study); "
        "OS: overall survival; OS60: 60-month truncated overall survival; "
        "HM450: Illumina HumanMethylation450; GEO: Gene Expression Omnibus; "
        "FDR: false discovery rate; BH: Benjamini–Hochberg; C-index: concordance index; "
        "RNA: ribonucleic acid; lncRNA: long non-coding RNA; GDC: Genomic Data Commons; "
        "EMT: epithelial–mesenchymal transition; IFN-γ: interferon gamma."
    )
    # insert after the new List heading: find it again
    lab = find_para(doc, "List of abbreviations")
    insert_after(lab, abbr, "Normal")

    # 8) Declarations umbrella + BMC subheadings (existing sentences only)
    insert_before(find_para(doc, "Acknowledgements"), "Declarations", "Heading 1")

    # Funding pulled from existing Acknowledgements sentence
    ack = find_para(doc, "Acknowledgements")
    ack_body = ack._element.getnext()
    ack_text = ptext(ack_body)
    funding_sent = "Funding: [insert grant agency and award number(s) before submission]."
    if funding_sent not in ack_text:
        raise RuntimeError("Expected funding placeholder in Acknowledgements")
    new_ack = ack_text.replace(funding_sent, "").replace("  ", " ").strip()
    ack_para = Paragraph(ack_body, ack._parent)
    set_highlighted(ack_para, new_ack)

    # Rename remaining declaration headings to BMC labels
    set_highlighted(find_para(doc, "Author contributions"), "Authors' contributions")
    set_highlighted(find_para(doc, "Data availability"), "Availability of data and materials")
    # Drop the extra Code availability heading; keep its paragraph under data availability.
    code_h = find_para(doc, "Code availability")
    code_h._element.getparent().remove(code_h._element)
    # Split ethics into two BMC headings using existing clauses.
    eth = find_para(doc, "Ethics declarations")
    eth_body_el = eth._element.getnext()
    eth_text = ptext(eth_body_el)
    expected_eth = (
        "Ethics approval and consent to participate: Not applicable. "
        "This study re-analysed de-identified, publicly available TCGA and GEO data with no access to identifiable participant information. "
        "Consent for publication: Not applicable."
    )
    if eth_text != expected_eth:
        raise RuntimeError(f"Ethics text changed unexpectedly:\n{eth_text}")
    set_highlighted(eth, "Ethics approval and consent to participate")
    set_highlighted(
        Paragraph(eth_body_el, eth._parent),
        "Not applicable. This study re-analysed de-identified, publicly available TCGA and GEO data with no access to identifiable participant information.",
    )
    p_pub_h = insert_after(Paragraph(eth_body_el, eth._parent), "Consent for publication", "Heading 1")
    insert_after(p_pub_h, "Not applicable.", "Normal")

    # Insert Funding heading + existing placeholder after Acknowledgements body
    ack = find_para(doc, "Acknowledgements")
    ack_body_p = Paragraph(ack._element.getnext(), ack._parent)
    fund_h = insert_after(ack_body_p, "Funding", "Heading 1")
    insert_after(fund_h, "[insert grant agency and award number(s) before submission].", "Normal")

    # Competing interests heading stays; highlight as BMC-required declaration item
    set_highlighted(find_para(doc, "Competing interests"), "Competing interests")

    # Place abbreviations, Declarations (BMC subheading order), then References before Figure legends
    h1_els = heading1_elements(doc)
    fig_el = find_para(doc, "Figure legends")._element
    ordered_blocks = [
        collect_until_next_h1(find_para(doc, "List of abbreviations")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Declarations")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Ethics approval and consent to participate")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Consent for publication")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Availability of data and materials")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Competing interests")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Funding")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Authors' contributions")._element, h1_els),
        collect_until_next_h1(find_para(doc, "Acknowledgements")._element, h1_els),
        collect_until_next_h1(find_para(doc, "References")._element, h1_els),
    ]
    for block in ordered_blocks:
        for el in block:
            fig_el.addprevious(el)

    # 9) Vancouver punctuation on numbered references (content of each citation unchanged)
    in_refs = False
    for p in doc.paragraphs:
        t = p.text.strip()
        if t == "References":
            in_refs = True
            continue
        if in_refs and p.style.name.startswith("Heading"):
            break
        if in_refs and re.match(r"^\d+\.\s", t):
            new = vancouverize_original(t)
            if new != t:
                set_highlighted(p, new)

    out = ROOT / "Manuscript First.docx"
    doc.save(str(out))
    print("Backup:", bak)
    print("Wrote:", out)

    # verify order
    heads = [p.text.strip() for p in Document(str(out)).paragraphs if p.style.name.startswith("Heading 1")]
    print("Heading 1 order:")
    for h in heads:
        print(" -", h)


if __name__ == "__main__":
    main()
