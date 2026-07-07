## TCGA ESCA/ESCC — code package (Scientific Reports submission)

Bu klasör (`code/repo/`), bu manuscript için ürettiğimiz **kodları** GitHub’a koymaya hazır şekilde paketler.

### Klasör yapısı

- `esca_escc_multiomic_pipeline/`: ana TCGA analizi + figür/tablo üretimi + manuscript düzeltme scriptleri
- `geo_gse536xx_validation/`: GEO modülleri (scriptler + küçük raporlar)
- `submission_prep/`: submission hazırlık notları / checklist / dokümanlar
- `reference_verification_report.json`: referans doğrulama raporu
- `scientific_reports_format_check_report.txt`: SR format kontrol raporu

### Bilerek dahil edilmeyen (GitHub için çok büyük / redistribüsyon uygun değil)

- **Ham TCGA Firehose Legacy dosyaları** (`data_*.txt`, `meta_*.txt`, `case_lists/`)
- `result/`, `Plots/`, `Supplementary/` (çıktılar ve büyük dosyalar)
- GEO ham dosyaları (`*.xlsx`, `*.soft.gz`)
- Manuscript `.docx` ve `.bak` yedekleri

Bu dosyalar için “Data availability” kısmındaki kaynaklardan yeniden indirmeniz gerekir.

### Çalıştırma (özet)

Proje kökünden (bu repo kökü) çalıştırın:

```bash
# ana TCGA pipeline (modeller + etkileşimler + mediation vb.)
python3 -m esca_escc_multiomic_pipeline.run_all

# fast-death sensitivity → Supplementary Table S10
python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity

# figürler
python3 -m esca_escc_multiomic_pipeline.make_plots
python3 -m esca_escc_multiomic_pipeline.make_methylation_os_plots

# GEO modülleri (ham GEO dosyaları gerekiyorsa)
python3 geo_gse536xx_validation/run_combined_validation_161533_53624.py
```

> Not: `result/` ve `Plots/` gibi çıktı klasörleri scriptler tarafından tekrar üretilecektir (ham veriler mevcutsa).

