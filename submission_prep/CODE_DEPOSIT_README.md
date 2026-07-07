# TCGA ESCA prognostic transfer — code deposit guide

Use this folder structure when creating a **public GitHub repository** (and optional Zenodo archive) before Scientific Reports submission.

## Recommended repository name

`tcga-esca-prognostic-transfer`

## Minimum files to include

```
README.md                          ← copy from this guide + run instructions
requirements.txt                   ← see below
esca_escc_multiomic_pipeline/      ← entire package
geo_gse536xx_validation/           ← GEO scripts + reports (exclude large .soft.gz if needed)
gdc2025_esca_reanalysis.py         ← GDC cross-validation
result/                            ← processed TSV outputs (or Zenodo-only if repo size limits)
Plots/                             ← figure PNGs (optional; can be Zenodo-only)
```

## requirements.txt (suggested)

```
pandas>=2.0
numpy>=1.24
scipy>=1.10
lifelines>=0.27
matplotlib>=3.7
seaborn>=0.13
python-docx>=1.0
mygene>=3.2
```

## Reproduction commands

From the study root (parent of `esca_escc_multiomic_pipeline/`):

```bash
# Primary TCGA models, mediation, mutation interactions
python3 -m esca_escc_multiomic_pipeline.run_all

# Fast-death sensitivity → Supplementary Table S10
python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity

# Figures
export MPLCONFIGDIR=/tmp/mpl && mkdir -p "$MPLCONFIGDIR"
python3 -m esca_escc_multiomic_pipeline.make_methylation_os_plots
python3 -m esca_escc_multiomic_pipeline.make_plots

# GEO validation (requires GEO matrix files in geo_gse536xx_validation/)
python3 geo_gse536xx_validation/run_combined_validation_161533_53624.py

# GDC 2025 cross-validation (requires separate GDC download folder)
python3 gdc2025_esca_reanalysis.py
```

## Raw TCGA data

Not redistributed in the repo. Users download the ESCA Firehose Legacy study package from cBioPortal and place TSV files in the study root (`data_clinical_*.txt`, `data_methylation_hm450.txt`, etc.).

## Zenodo

After creating the GitHub release, link it on Zenodo for a citable DOI and replace `[INSERT ZENODO DOI]` in the manuscript Data availability paragraph.

## Placeholder to replace in manuscript

- `[INSERT PUBLIC REPOSITORY URL]` → your GitHub URL
- `[INSERT ZENODO DOI ON FIRST SUBMISSION]` → `https://doi.org/10.5281/zenodo.XXXXXXX`
