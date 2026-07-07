"""
Renumber manuscript tables: 2A→2, 2B→3, 2C→4, 2D→5; former Table 3 (Cox)→6; former Table 4 (mechanistic)→7.
Updates every paragraph in the main manuscript docx (idempotent if run twice with care).
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document


def renumber_table_text(t: str) -> str:
    if not t:
        return t

    # --- Escalate former Table 4 (mechanistic) → 7 first ---
    t = t.replace(
        "Table 4. Exploratory mechanistic follow-up:",
        "Table 7. Exploratory mechanistic follow-up:",
    )
    t = re.sub(
        r"\bTable 4 lists the strongest mutation",
        "Table 7 lists the strongest mutation",
        t,
    )
    t = re.sub(r"\bFigure 5 and Table 4 supply", "Figure 5 and Table 7 supply", t)
    t = re.sub(r"\bFigure 6 and Table 4 as exploratory", "Figure 6 and Table 7 as exploratory", t)
    t = re.sub(r"\bpositioned Figure 6 and Table 4", "positioned Figure 6 and Table 7", t)
    t = re.sub(r"\(Figures 5–6; Table 4\)", "(Figures 5–6; Table 7)", t)
    t = re.sub(r"Table 4\) motivates", "Table 7) motivates", t)
    t = re.sub(r"Table 4\) frames", "Table 7) frames", t)
    t = re.sub(r"Table 4 build on", "Table 7 build on", t)
    t = re.sub(r"and Table 4\.", "and Table 7.", t)
    t = re.sub(r"See Results \(mutation subsection\) and Table 4\.", "See Results (mutation subsection) and Table 7.", t)

    # --- Former Cox Table 3 → 6 ---
    t = t.replace(
        "Table 3. EAC/ESCA: univariate Cox associations",
        "Table 6. EAC/ESCA: univariate Cox associations",
    )
    t = t.replace("(Figure 4; Table 3)", "(Figure 4; Table 6)")
    t = re.sub(r"\(Table 3\)\. Kaplan", "(Table 6). Kaplan", t)
    t = re.sub(r"associations \(Table 3\)\.", "associations (Table 6).", t)
    t = re.sub(r"ranking in Table 3 to", "ranking in Table 6 to", t)
    t = re.sub(r"Figure 4 and Table 3 demonstrate", "Figure 4 and Table 6 demonstrate", t)
    t = re.sub(r"\(Figure 4; Table 3\)", "(Figure 4; Table 6)", t)
    t = re.sub(r"Table 3 summarizes EAC-only", "Table 6 summarizes EAC-only", t)
    t = re.sub(r"guided by Table 3\.", "guided by Table 6.", t)
    t = re.sub(r"with EAC priorities guided by Table 3", "with EAC priorities guided by Table 6", t)
    t = re.sub(r"separation \(Figure 4; Table 3\)", "separation (Figure 4; Table 6)", t)
    t = re.sub(r"and Table 3\)", "and Table 6)", t)
    t = re.sub(
        r"Figure 4 and Table 3 use the same",
        "Figure 4 and Table 6 use the same",
        t,
    )
    t = re.sub(r"to univariate Cox rankings \(Table 3\)", "to univariate Cox rankings (Table 6)", t)
    t = re.sub(r"Methylation–OS associations in EAC \(Table 3\)", "Methylation–OS associations in EAC (Table 6)", t)
    t = re.sub(r"same clinical horizon \(Figure 4; Table 3\)", "same clinical horizon (Figure 4; Table 6)", t)
    t = re.sub(r"and Table 3\.", "and Table 6.", t)
    t = re.sub(r"and Table 3,", "and Table 6,", t)
    t = re.sub(r"See Results \(KM subsection\) and Table 3\.", "See Results (KM subsection) and Table 6.", t)

    # --- Ranges spanning all tables ---
    t = t.replace("Tables 1–4", "Tables 1–7")
    t = t.replace("Tables 1–4,", "Tables 1–7,")
    t = t.replace("Tables 1–4.", "Tables 1–7.")
    t = t.replace("Tables 1–4 should", "Tables 1–7 should")

    # --- Composite 2A–2D (before splitting A/B/C/D) ---
    t = t.replace("Tables 2A–2D", "Tables 2–5")
    t = t.replace("Tables 2A–2D,", "Tables 2–5,")
    t = t.replace("Tables 2A–2D.", "Tables 2–5.")
    t = t.replace("Tables 2A–2D—", "Tables 2–5—")
    t = t.replace("Tables 2A–2D:", "Tables 2–5:")
    t = t.replace("Tables 2A and 2B", "Tables 2 and 3")
    t = t.replace("Tables 2C and 2D", "Tables 4 and 5")
    t = t.replace("Tables 2A–2B", "Tables 2–3")
    t = t.replace("Tables 2A–2B)", "Tables 2–3)")

    # --- Captions 2D→5 … 2A→2 (reverse order) ---
    t = t.replace(
        "Table 2D. ESCC: top hypermethylated HM450 probes",
        "Table 5. ESCC: top hypermethylated HM450 probes",
    )
    t = t.replace(
        "Table 2C. EAC/ESCA: top hypermethylated HM450 probes",
        "Table 4. EAC/ESCA: top hypermethylated HM450 probes",
    )
    t = t.replace(
        "Table 2B. ESCC: top hypomethylated HM450 probes",
        "Table 3. ESCC: top hypomethylated HM450 probes",
    )
    t = t.replace(
        "Table 2A. EAC/ESCA: top hypomethylated HM450 probes",
        "Table 2. EAC/ESCA: top hypomethylated HM450 probes",
    )

    # --- Inline citations (gene lists) ---
    t = re.sub(r"\(Table 2A\)", "(Table 2)", t)
    t = re.sub(r"\(Table 2B\)", "(Table 3)", t)
    t = re.sub(r"\(Table 2C\)", "(Table 4)", t)
    t = re.sub(r"\(Table 2D\)", "(Table 5)", t)
    t = re.sub(r"Table 2A\)", "Table 2)", t)
    t = re.sub(r"Table 2B\)", "Table 3)", t)
    t = re.sub(r"Table 2C\)", "Table 4)", t)
    t = re.sub(r"Table 2D\)", "Table 5)", t)
    t = re.sub(r"Table 2A,", "Table 2,", t)
    t = re.sub(r"Table 2B,", "Table 3,", t)
    t = re.sub(r"Table 2C,", "Table 4,", t)
    t = re.sub(r"Table 2D,", "Table 5,", t)
    t = re.sub(r"in EAC \(Table 2A\)", "in EAC (Table 2)", t)
    t = re.sub(r"among ESCC leaders \(Table 2B\)", "among ESCC leaders (Table 3)", t)
    t = re.sub(r"ranks \(Table 2C\)", "ranks (Table 4)", t)
    t = re.sub(r"in Table 2D\.", "in Table 5.", t)
    t = re.sub(r"hypermethylated ranks in Table 2D", "hypermethylated ranks in Table 5", t)
    t = re.sub(r"with complementary hypermethylated ranks \(Table 2C\)", "with complementary hypermethylated ranks (Table 4)", t)

    # Abstract / intro arcs
    t = t.replace("(Figure 4; Table 3); and (iv)", "(Figure 4; Table 6); and (iv)")
    t = t.replace("(Figures 5–6; Table 4).", "(Figures 5–6; Table 7).")
    t = t.replace("(Figures 5–6; Table 4)", "(Figures 5–6; Table 7)")
    t = t.replace("Primary numeric summaries appear in Tables 1–4;", "Primary numeric summaries appear in Tables 1–7;")
    t = t.replace("appear in Figures 1–6.", "appear in Figures 1–6.")  # no-op

    t = re.sub(
        r"Figures 2–3 and Tables 2A–2D derive from",
        "Figures 2–3 and Tables 2–5 derive from",
        t,
    )
    t = re.sub(
        r"Figures 5–6 and Table 4 build on",
        "Figures 5–6 and Table 7 build on",
        t,
    )

    # Methods roadmap
    t = t.replace("(Figures 2–3; Tables 2A–2D); (iii)", "(Figures 2–3; Tables 2–5); (iii)")
    t = t.replace("(Figure 4; Table 3); and (iv)", "(Figure 4; Table 6); and (iv)")

    # Figure legends
    t = t.replace("and Tables 2A–2D.", "and Tables 2–5.")

    # Mechanistic layer: old "Table 4" in prose collided with new Table 4 (EAC hyper caption).
    t = t.replace(
        "mutation-linked module shifts (Figure 5; Table 4) and",
        "mutation-linked module shifts (Figure 5; Table 7) and",
    )
    t = t.replace("Figure 5 and Table 4 offer", "Figure 5 and Table 7 offer")

    return t


def main():
    study = Path(__file__).resolve().parents[1]
    path = study / "TCGA_ESCA_ESCC_landmark60_full_scientific_manuscript.docx"
    doc = Document(str(path))
    for p in doc.paragraphs:
        p.text = renumber_table_text(p.text)
    doc.save(str(path))
    print("Table numbering updated:", path)


if __name__ == "__main__":
    main()
