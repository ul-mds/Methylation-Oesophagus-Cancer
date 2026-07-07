"""
Round numeric tokens in all Word tables of the manuscript to three digits after
the decimal point. Values with |x| < 0.001 (non-zero) are shown in scientific
notation with three decimals in the mantissa (e.g. 4.326e-06) so q-values are
not collapsed to 0.000. Integer tokens (no decimal point) are left unchanged.
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document

# Floats with optional scientific suffix; integers without '.' are not matched.
_FLOAT_RE = re.compile(
    r"-?\d+\.\d+(?:[eE][+-]?\d+)?|" r"-?\d+[eE][+-]?\d+"
)


def _format_float(x: float) -> str:
    if x != x:  # NaN
        return "nan"
    if x == 0:
        return "0.000"
    ax = abs(x)
    if ax >= 1e-3:
        return f"{x:.3f}"
    return f"{x:.3e}"


def _round_numeric_tokens(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        return _format_float(float(match.group(0)))

    return _FLOAT_RE.sub(repl, text)


def main():
    path = Path(__file__).resolve().parents[1] / "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"
    doc = Document(str(path))
    n_cells = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                old = cell.text
                new = _round_numeric_tokens(old)
                if new != old:
                    cell.text = new
                    n_cells += 1
    doc.save(str(path))
    print(f"Updated {n_cells} table cells with rounded numerics:", path)


if __name__ == "__main__":
    main()
