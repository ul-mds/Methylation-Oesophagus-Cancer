"""Remove the gene_token column from Word tables in the manuscript (idempotent)."""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def _delete_table_column(table, col_idx: int) -> None:
    for row in table.rows:
        tr = row._element
        tcs = tr.findall(qn("w:tc"))
        if col_idx < len(tcs):
            tr.remove(tcs[col_idx])


def main():
    path = Path(__file__).resolve().parents[1] / "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"
    doc = Document(str(path))
    n = 0
    for table in doc.tables:
        if not table.rows:
            continue
        hdr = [c.text.strip() for c in table.rows[0].cells]
        if len(hdr) > 1 and hdr[1] == "gene_token":
            _delete_table_column(table, 1)
            n += 1
    doc.save(str(path))
    print(f"Removed gene_token column from {n} tables:", path)


if __name__ == "__main__":
    main()
