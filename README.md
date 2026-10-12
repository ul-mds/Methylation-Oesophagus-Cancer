## TCGA ESCA/ESCC — analysis code (Cancer Medicine)

Public code companion for the **Cancer Medicine** manuscript on histology-specific methylation and prognostic transfer in TCGA esophageal carcinoma (ESCA vs ESCC).

**Repository:** https://github.com/ul-mds/Methylation-Oesophagus-Cancer  
**Author guidelines:** https://onlinelibrary.wiley.com/page/journal/20457634/homepage/forauthors.html

### Layout

- `esca_escc_multiomic_pipeline/` — TCGA models, figures, tables, and optional manuscript formatting utilities
- `geo_gse536xx_validation/` — GEO validation scripts
- `submission_prep/` — checklists and submission notes
- `cancer_medicine_format_check_report.txt` — format checklist snapshot (legacy manuscript filename)

### Not included (download separately)

- Raw TCGA Firehose Legacy TSVs (`data_*.txt`, `meta_*.txt`, `case_lists/`)
- Large outputs (`result/`, `Plots/`, supplementary bundles)
- Manuscript `.docx` files (see local `Cancer Medicine submission_package`)

### Reproduction

From this repository root (with TCGA files in `ESCA_STUDY_ROOT` or current directory as documented in each script):

```bash
pip install -r requirements.txt

python3 -m esca_escc_multiomic_pipeline.run_all
python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity
python3 -m esca_escc_multiomic_pipeline.make_plots
python3 -m esca_escc_multiomic_pipeline.make_methylation_os_plots
python3 geo_gse536xx_validation/run_combined_validation_161533_53624.py
```

For manuscript packaging scripts that target `Cancer Medicine submission_package`, set:

```bash
export ESCA_STUDY_ROOT="/path/to/Esophageal Carcinoma (TCGA, Firehose Legacy)"
export CM_SUBMISSION_PACKAGE="$ESCA_STUDY_ROOT/Cancer Medicine submission_package"
```

Figures use **Times New Roman** styling per Cancer Medicine figure preparation (`plot_style.py`).
