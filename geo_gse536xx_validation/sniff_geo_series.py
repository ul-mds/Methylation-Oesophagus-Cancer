"""Download GEO series matrix headers only (first ~400 lines) and summarize for validation checklist."""
from __future__ import annotations

import gzip
import re
import subprocess
import sys
from pathlib import Path
from urllib.request import urlopen

GSES = [
    "GSE23400",
    "GSE121930",
    "GSE52826",
    "GSE38129",
    "GSE20347",
    "GSE17351",
    "GSE45670",
    "GSE161533",
    "GSE51287",
    "GSE77861",
    "GSE53624",
    "GSE53622",
]

BASE = "https://ftp.ncbi.nlm.nih.gov/geo/series"


def folder_for(acc: str) -> str:
    n = acc.replace("GSE", "")
    pre = n[: len(n) - 3]
    return f"GSE{pre}nnn"


def list_matrix_name(acc: str) -> str | None:
    import urllib.request

    url = f"{BASE}/{folder_for(acc)}/{acc}/matrix/"
    try:
        html = urllib.request.urlopen(url, timeout=60).read().decode("utf-8", errors="replace")
    except Exception:
        return None
    m = re.findall(r'href="([^"]+_series_matrix\.txt\.gz)"', html)
    return m[0] if m else None


def fetch_header_gzip(url: str, max_lines: int = 500) -> str:
    # curl supports byte range poorly on this server; stream with subprocess
    proc = subprocess.Popen(
        ["curl", "-sL", "--max-time", "120", url],
        stdout=subprocess.PIPE,
    )
    assert proc.stdout
    dec = gzip.GzipFile(fileobj=proc.stdout)
    lines: list[str] = []
    for i, raw in enumerate(dec):
        if i >= max_lines:
            proc.kill()
            break
        try:
            lines.append(raw.decode("utf-8", errors="replace").rstrip("\n"))
        except Exception:
            pass
    return "\n".join(lines)


def parse_header(text: str) -> dict:
    title = ""
    platforms: list[str] = []
    summary_chunks: list[str] = []
    n_samples = None
    surv_hits: list[str] = []
    meth_hint = False
    rna_hint = False
    for line in text.splitlines():
        if line.startswith("!Series_title"):
            title = line.split("\t", 1)[-1].strip().strip('"')
        if line.startswith("!Series_platform_id"):
            platforms.append(line.split("\t", 1)[-1].strip().strip('"'))
        if line.startswith("!Series_summary"):
            summary_chunks.append(line.split("\t", 1)[-1].strip().strip('"'))
        if line.startswith("!Sample_geo_accession"):
            n_samples = len(line.split("\t")) - 1
        low = line.lower()
        if line.startswith("!Sample_characteristics") or line.startswith("!Sample_molecule_ch1"):
            if any(
                k in low
                for k in (
                    "survival",
                    "alive",
                    "death",
                    "follow",
                    "os:",
                    "dfs",
                    "prognos",
                    "vital",
                    "recurrence",
                    "month",
                )
            ):
                surv_hits.append(line[:300])
            if "methyl" in low:
                meth_hint = True
            if "rna" in low or "transcript" in low or "mrna" in low:
                rna_hint = True
        if "methylation" in low or "450" in low or "illumin" in low and "meth" in low:
            meth_hint = True
        if "expression" in low or "affymetrix" in low or "rna-seq" in low or "transcript" in low:
            rna_hint = True
    return {
        "title": title,
        "platforms": list(dict.fromkeys(platforms)),
        "summary": " ".join(summary_chunks)[:600],
        "n_samples": n_samples,
        "surv_lines": surv_hits[:5],
        "meth_hint": meth_hint,
        "rna_hint": rna_hint,
    }


def main() -> None:
    out_lines: list[str] = []
    for acc in GSES:
        mat = list_matrix_name(acc)
        if not mat:
            out_lines.append(f"{acc}\tNO_MATRIX\t-\t-\t-\n")
            continue
        url = f"{BASE}/{folder_for(acc)}/{acc}/matrix/{mat}"
        try:
            hdr = fetch_header_gzip(url, max_lines=600)
        except Exception as e:
            out_lines.append(f"{acc}\tFETCH_FAIL\t{mat}\t{e}\t\n")
            continue
        p = parse_header(hdr)
        surv = "YES" if p["surv_lines"] else "UNKNOWN"
        if p["surv_lines"]:
            surv += ": " + p["surv_lines"][0][:120].replace("\t", " ")
        out_lines.append(
            "\t".join(
                [
                    acc,
                    mat,
                    "|".join(p["platforms"]) or "?",
                    str(p["n_samples"] or "?"),
                    p["title"][:120].replace("\t", " "),
                    surv,
                    "meth_meta" if p["meth_hint"] else "",
                    "rna_meta" if p["rna_hint"] else "",
                ]
            )
            + "\n"
        )
    Path(__file__).resolve().parent.joinpath("geo_checklist_raw.tsv").write_text(
        "acc\tmatrix\tgpl\tn_samples\ttitle\tsurvival_sniff\tmeth\trna\n" + "".join(out_lines),
        encoding="utf-8",
    )
    print(Path(__file__).resolve().parent.joinpath("geo_checklist_raw.tsv").read_text())


if __name__ == "__main__":
    main()
