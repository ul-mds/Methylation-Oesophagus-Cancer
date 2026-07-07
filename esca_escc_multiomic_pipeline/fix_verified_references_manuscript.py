"""
Replace Manuscript First.docx reference list with PubMed/Crossref-verified entries
(Nature / Scientific Reports style: italic journal, bold volume, DOI included).
"""
from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph

STUDY_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DOCX = STUDY_ROOT / "Manuscript First.docx"

# (prefix before journal, journal, volume or None, suffix after volume)
VERIFIED_REFS: list[tuple[str, str, str | None, str]] = [
    (
        "Thrift, A. P. The epidemic of oesophageal carcinoma: where are we now? ",
        "Cancer Epidemiol.",
        "41",
        ", 88-95 (2016); https://doi.org/10.1016/j.canep.2016.01.013.",
    ),
    (
        "Spechler, S. J. & Souza, R. F. Barrett's esophagus. ",
        "N. Engl. J. Med.",
        "371",
        ", 836-845 (2014); https://doi.org/10.1056/NEJMra1314704.",
    ),
    (
        "Cancer Genome Atlas Research Network. Integrated genomic characterization of oesophageal carcinoma. ",
        "Nature",
        "541",
        ", 169-175 (2017); https://doi.org/10.1038/nature20805.",
    ),
    (
        "Lyu, Q. et al. Integrated bioinformatics analysis of differences between EAC and ESCC. ",
        "BMC Cancer",
        "25",
        ", 1668 (2025); https://doi.org/10.1186/s12885-025-15090-z.",
    ),
    (
        "Pinto, R. et al. Molecular stratification of esophageal adenocarcinoma: implications for prognosis and treatment strategy. ",
        "Oncogene",
        "45",
        ", 353-367 (2026); https://doi.org/10.1038/s41388-025-03650-3.",
    ),
    (
        "Cui, H. et al. The consensus molecular subtypes of esophageal squamous cell carcinoma. ",
        "Sig. Transduct. Target. Ther.",
        "11",
        ", 65 (2026); https://doi.org/10.1038/s41392-026-02577-9.",
    ),
    (
        "Zheng, Y. et al. Genome-wide DNA methylation and gene expression profiling characterizes molecular subtypes of esophagus squamous cell carcinoma for predicting patient survival and immunotherapy efficacy. ",
        "Cancers",
        "14",
        ", 4970 (2022); https://doi.org/10.3390/cancers14204970.",
    ),
    (
        "van de Velde, C. J. H. et al. Preoperative chemoradiotherapy for esophageal or junctional cancer. ",
        "N. Engl. J. Med.",
        "366",
        ", 2074-2084 (2012); https://doi.org/10.1056/NEJMoa1112087.",
    ),
    (
        "Bremner, C. G., Lynch, V. P. & Ellis, F. H. Jr. Barrett's esophagus: congenital or acquired? ",
        "Surgery",
        "68",
        ", 209-216 (1970); https://doi.org/10.1016/0039-6060(70)90326-5.",
    ),
    (
        "Tosh, D. & Slack, J. M. How cells change their phenotype. ",
        "Nat. Rev. Mol. Cell Biol.",
        "3",
        ", 187-194 (2002); https://doi.org/10.1038/nrm761.",
    ),
    (
        "Sarosi, G. et al. Bone marrow progenitor cells contribute to esophageal regeneration and metaplasia in a rat model of Barrett's esophagus. ",
        "Dis. Esophagus",
        "21",
        ", 43-50 (2008); https://doi.org/10.1111/j.1442-2050.2007.00744.x.",
    ),
    (
        "Hu, Q. et al. Using Barrett's esophagus as a model to study stem cell differentiation. ",
        "J. Gastroenterol.",
        "49",
        ", 287-294 (2014); https://doi.org/10.1007/s00535-013-0883-4.",
    ),
    (
        "Kunene, V. et al. Prognostic markers in oesophageal and gastric cancer review. Are they ready for clinical practice? ",
        "ESMO Gastrointest. Oncol.",
        "1",
        ", 100091 (2024); https://doi.org/10.1016/j.esmogo.2024.100091.",
    ),
    (
        "Mandard, A. M. et al. Pathologic assessment of tumor regression after preoperative chemoradiotherapy. ",
        "Int. J. Radiat. Oncol. Biol. Phys.",
        "29",
        ", 631-633 (1994); https://doi.org/10.1016/0360-3016(94)00433-4.",
    ),
    (
        "Zhang, X. et al. Multi-omics analysis unveils a four-gene prognostic signature in esophageal squamous carcinoma and the therapeutic potential of PKP1. ",
        "BMC Cancer",
        "25",
        ", 777 (2025); https://doi.org/10.1186/s12885-025-14150-8.",
    ),
    (
        "Oesophageal Cancer Clinical and Molecular Stratification Consortium et al. Disentangling oncogenic amplicons in esophageal adenocarcinoma. ",
        "Nat. Commun.",
        "15",
        ", 4074 (2024); https://doi.org/10.1038/s41467-024-47619-4.",
    ),
    (
        "Chang, Z. et al. Construction of a multimodal esophageal cancer prognostic model with screening of key genes and investigation of their biological functions. ",
        "Discov. Oncol.",
        "16",
        ", 1560 (2025); https://doi.org/10.1007/s12672-025-03404-3.",
    ),
    (
        "Tang, P. et al. Integrated machine learning developed a prognosis-related gene signature to predict prognosis in oesophageal squamous cell carcinoma. ",
        "J. Cell. Mol. Med.",
        "28",
        ", e70171 (2024); https://doi.org/10.1111/jcmm.70171.",
    ),
    (
        "Tuersun, A. et al. Interpretable machine learning models to predict survival in esophageal cancer: a study based on the SEER database and external validation in China. ",
        "Front. Physiol.",
        "16",
        ", 1665383 (2025); https://doi.org/10.3389/fphys.2025.1665383.",
    ),
    (
        "Pu, Y. et al. Estimating the prognosis of esophageal squamous cell carcinoma based on The Cancer Genome Atlas (TCGA) of m6A methylation-associated genes. ",
        "J. Gastrointest. Oncol.",
        "13",
        ", 1-12 (2022); https://doi.org/10.21037/jgo-21-686.",
    ),
    (
        "Lu, T. et al. Identification of DNA methylation-driven genes in esophageal squamous cell carcinoma: a study based on The Cancer Genome Atlas. ",
        "Cancer Cell Int.",
        "19",
        ", 52 (2019); https://doi.org/10.1186/s12935-019-0770-9.",
    ),
    (
        "Chen, Y. et al. Identification of prognostic risk model based on DNA methylation-driven genes in esophageal adenocarcinoma. ",
        "Biomed. Res. Int.",
        "2021",
        ", 6628391 (2021); https://doi.org/10.1155/2021/6628391.",
    ),
    (
        "Cao, W. et al. Multi-faceted epigenetic dysregulation of gene expression promotes esophageal squamous cell carcinoma. ",
        "Nat. Commun.",
        "11",
        ", 3675 (2020); https://doi.org/10.1038/s41467-020-19613-8.",
    ),
    (
        "Xi, Y. et al. Multi-omic characterization of genome-wide abnormal DNA methylation reveals diagnostic and prognostic markers for esophageal squamous-cell carcinoma. ",
        "Sig. Transduct. Target. Ther.",
        "7",
        ", 53 (2022); https://doi.org/10.1038/s41392-022-00953-0.",
    ),
    (
        "Zhao, D. et al. Advances in multi-omics for esophageal squamous cell carcinoma: diagnostic, prognostic, and therapeutic perspectives. ",
        "Protein Cell",
        None,
        " (2026); https://doi.org/10.1093/procel/pwag005.",
    ),
    (
        "Xu, Y. et al. DNA methylation markers in esophageal cancer. ",
        "Front. Genet.",
        "15",
        ", 1354195 (2024); https://doi.org/10.3389/fgene.2024.1354195.",
    ),
    (
        "Tang, J. et al. DNA methylation in esophageal cancer: technological advances and early detection clinical applications. ",
        "Front. Oncol.",
        "15",
        ", 1543190 (2025); https://doi.org/10.3389/fonc.2025.1543190.",
    ),
    (
        "Wang, G. et al. Analysis of the associations between DNA methylation and clinical features reveals the potential oncogenic effect of CFAP52 in esophageal squamous cell cancer. ",
        "Cancer Cell Int.",
        "25",
        ", 394 (2025); https://doi.org/10.1186/s12935-025-03934-4.",
    ),
    (
        "Carter, S. L. et al. Absolute quantification of somatic DNA alterations in human cancer. ",
        "Nat. Biotechnol.",
        "30",
        ", 413-421 (2012); https://doi.org/10.1038/nbt.2203.",
    ),
    (
        "Liberzon, A. et al. The Molecular Signatures Database (MSigDB) hallmark gene set collection. ",
        "Cell Syst.",
        "1",
        ", 417-425 (2015); https://doi.org/10.1016/j.cels.2015.12.004.",
    ),
]


def _write_ref(p: Paragraph, num: int, prefix: str, journal: str, volume: str | None, suffix: str) -> None:
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
                if idx >= len(VERIFIED_REFS):
                    raise RuntimeError(f"More refs than VERIFIED_REFS ({idx + 1})")
                prefix, journal, volume, suffix = VERIFIED_REFS[idx]
                _write_ref(p, idx + 1, prefix, journal, volume, suffix)
                idx += 1
    if idx != len(VERIFIED_REFS):
        raise RuntimeError(f"Expected {len(VERIFIED_REFS)} references, rewrote {idx}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", type=Path, default=DEFAULT_DOCX)
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()
    path = args.docx.resolve()
    if not path.is_file():
        raise SystemExit(f"Not found: {path}")

    if not args.no_backup:
        bak = path.with_suffix(path.suffix + ".pre_verified_refs.bak")
        shutil.copy2(path, bak)
        print(f"Backup: {bak}")

    doc = Document(str(path))
    _rewrite_reference_list(doc)
    doc.save(str(path))
    print(f"Updated {len(VERIFIED_REFS)} verified references with DOIs: {path}")


if __name__ == "__main__":
    main()
