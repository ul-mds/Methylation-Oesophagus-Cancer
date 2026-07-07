from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd


# Package lives inside the cBioPortal study folder; parent of this package is the study root.
ESCA_DIR = Path(__file__).resolve().parents[1]


def _read_cbio_tsv(path: Path) -> pd.DataFrame:
    """
    cBioPortal clinical files have comment/header metadata rows starting with '#',
    followed by the real header row.
    """
    rows: list[list[str]] = []
    header: list[str] | None = None
    with path.open("r", encoding="utf-8", errors="replace") as f:
        for line in f:
            if line.startswith("#"):
                continue
            parts = line.rstrip("\n").split("\t")
            if header is None:
                header = parts
                continue
            rows.append(parts)
    if header is None:
        raise ValueError(f"No header found in {path}")
    return pd.DataFrame(rows, columns=header)


def _na(series: pd.Series) -> pd.Series:
    return series.replace({"[Not Available]": pd.NA, "NA": pd.NA, "": pd.NA})


def read_clinical_patient(base_dir: Path = ESCA_DIR) -> pd.DataFrame:
    df = _read_cbio_tsv(base_dir / "data_clinical_patient.txt")
    df["PATIENT_ID"] = df["PATIENT_ID"].astype(str)
    df["OS_MONTHS"] = pd.to_numeric(_na(df["OS_MONTHS"]), errors="coerce")
    df["DFS_MONTHS"] = pd.to_numeric(_na(df["DFS_MONTHS"]), errors="coerce")
    df["OS_EVENT"] = _na(df["OS_STATUS"]).astype(str).str.startswith("1").astype("int64")
    df["DFS_EVENT"] = _na(df["DFS_STATUS"]).astype(str).str.startswith("1").astype("int64")
    return df


def read_clinical_sample(base_dir: Path = ESCA_DIR) -> pd.DataFrame:
    df = _read_cbio_tsv(base_dir / "data_clinical_sample.txt")
    df["PATIENT_ID"] = df["PATIENT_ID"].astype(str)
    df["SAMPLE_ID"] = df["SAMPLE_ID"].astype(str)
    if "TMB_NONSYNONYMOUS" in df.columns:
        df["TMB_NONSYNONYMOUS"] = pd.to_numeric(_na(df["TMB_NONSYNONYMOUS"]), errors="coerce")
    return df


def read_expression_zscores(base_dir: Path = ESCA_DIR, which: str = "all_samples") -> pd.DataFrame:
    if which == "all_samples":
        path = base_dir / "data_mrna_seq_v2_rsem_zscores_ref_all_samples.txt"
    elif which == "diploid":
        path = base_dir / "data_mrna_seq_v2_rsem_zscores_ref_diploid_samples.txt"
    else:
        raise ValueError("which must be 'all_samples' or 'diploid'")

    df = pd.read_csv(path, sep="\t", low_memory=False)
    # First column is usually gene identifier (Hugo symbol)
    gene_col = df.columns[0]
    df = df.set_index(gene_col)
    df.index = df.index.astype(str)
    # Some exports contain duplicated gene symbols; collapse duplicates deterministically.
    if df.index.has_duplicates:
        df = df.groupby(level=0).mean(numeric_only=True)
    return df


def read_methylation(base_dir: Path = ESCA_DIR) -> pd.DataFrame:
    # HM450 beta values, rows=probes, cols=samples
    df = pd.read_csv(base_dir / "data_methylation_hm450.txt", sep="\t", low_memory=False)
    probe_col = df.columns[0]
    df = df.set_index(probe_col)
    df.index = df.index.astype(str)
    # Rarely, probe IDs can be duplicated in cBioPortal exports; collapse by mean.
    if df.index.has_duplicates:
        df = df.groupby(level=0).mean(numeric_only=True)
    return df


def read_mutations(base_dir: Path = ESCA_DIR) -> pd.DataFrame:
    # MAF-like with many columns; we only need Tumor_Sample_Barcode and Hugo_Symbol typically.
    df = pd.read_csv(base_dir / "data_mutations.txt", sep="\t", low_memory=False)
    # Normalize common column names
    if "Tumor_Sample_Barcode" not in df.columns:
        # some cbio exports use SAMPLE_ID
        for alt in ["SAMPLE_ID", "Sample_ID", "Tumor_Sample_Barcode"]:
            if alt in df.columns:
                df = df.rename(columns={alt: "Tumor_Sample_Barcode"})
                break
    if "Hugo_Symbol" not in df.columns:
        for alt in ["HUGO_SYMBOL", "Gene", "Hugo_Symbol"]:
            if alt in df.columns:
                df = df.rename(columns={alt: "Hugo_Symbol"})
                break
    df["Tumor_Sample_Barcode"] = df["Tumor_Sample_Barcode"].astype(str)
    df["Hugo_Symbol"] = df["Hugo_Symbol"].astype(str)
    return df


def build_master_clinical() -> pd.DataFrame:
    pat = read_clinical_patient()
    samp = read_clinical_sample()

    # One row per sample (not patient) because expression/methylation are sample-level.
    df = samp.merge(
        pat[
            [
                "PATIENT_ID",
                "OS_MONTHS",
                "OS_EVENT",
                "DFS_MONTHS",
                "DFS_EVENT",
                "AGE",
                "SEX",
            ]
        ],
        on="PATIENT_ID",
        how="left",
    )

    # Keep primary tumors and ESCA/ESCC histologies
    if "SAMPLE_TYPE" in df.columns:
        df = df[df["SAMPLE_TYPE"].astype(str).str.lower().str.contains("primary")]
    df = df[df["ONCOTREE_CODE"].isin(["ESCA", "ESCC"])].copy()
    df["HISTOLOGY"] = df["ONCOTREE_CODE"].map({"ESCA": "EAC/ESCA", "ESCC": "ESCC"})
    return df


def mutation_matrix(mutations: pd.DataFrame, samples: Iterable[str], genes: Iterable[str]) -> pd.DataFrame:
    genes = list(dict.fromkeys([str(g) for g in genes]))
    samples = list(dict.fromkeys([str(s) for s in samples]))
    df = mutations[mutations["Tumor_Sample_Barcode"].isin(samples) & mutations["Hugo_Symbol"].isin(genes)]
    mat = pd.crosstab(df["Tumor_Sample_Barcode"], df["Hugo_Symbol"]).reindex(index=samples, columns=genes, fill_value=0)
    mat = (mat > 0).astype("int64")
    mat.index.name = "SAMPLE_ID"
    return mat


def align_omics_to_samples(
    master: pd.DataFrame,
    expr: pd.DataFrame | None = None,
    meth: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None]:
    samples = master["SAMPLE_ID"].astype(str).tolist()

    def _subset_cols(mat: pd.DataFrame) -> pd.DataFrame:
        cols = [c for c in samples if c in mat.columns]
        return mat.loc[:, cols]

    expr2 = _subset_cols(expr) if expr is not None else None
    meth2 = _subset_cols(meth) if meth is not None else None

    # Further restrict master to samples present in all provided omics
    keep = set(samples)
    if expr2 is not None:
        keep &= set(expr2.columns)
    if meth2 is not None:
        keep &= set(meth2.columns)

    keep = sorted(keep)
    master2 = master[master["SAMPLE_ID"].isin(keep)].copy()
    if expr2 is not None:
        expr2 = expr2.loc[:, keep]
    if meth2 is not None:
        meth2 = meth2.loc[:, keep]
    return master2, expr2, meth2

