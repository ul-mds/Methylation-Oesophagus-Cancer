# Gönderim öncesi hazırlık raporu (Cover letter hariç)

**Tarih:** 21 Haziran 2026  
**Proje:** TCGA ESCA — histology-specific prognostic transfer  
**Manuscript güncellemesi:** Bu rapordaki metinleri siz `Manuscript.docx`'e aktaracaksınız.

---

## 1. Tamamlanan işler — özet

| # | Madde | Durum | Çıktı dosyası |
|---|--------|--------|----------------|
| 2 | Data availability | Taslak metin hazır | `submission_prep/data_availability_draft.txt` |
| 2b | Kod deposu rehberi | Hazır | `submission_prep/CODE_DEPOSIT_README.md` |
| 3 | REMARK checklist | Hazır | `submission_prep/remark_checklist.md` |
| 4 | Supplementary Table S10 | **Analiz çalıştırıldı** | `result/fast_death_sensitivity_*.tsv` + `submission_prep/supplementary_table_S10.md` |
| 5 | Response to Reviewer | Hazır (18 yorum) | `submission_prep/response_to_reviewer.md` |

**Yeni analiz scripti:** `esca_escc_multiomic_pipeline/run_fast_death_sensitivity.py`

---

## 2. Supplementary Table S10 — ana bulgular

### Panel A — Hasta sınıflandırması (min n=10 kuralı)

| Histology | %10 / %20 / %30 fast-death n | OS60 ölüm | Not |
|-----------|------------------------------|-----------|-----|
| **ESCC** | 10 / 10 / 10 | 32 | Üyelik üç eşikte **aynı** → Jaccard = 1.00 |
| **EAC** | 10 / 10 / **14** | 44 | %30'da 4 ek hasta |

### Panel B — Probe stabilitesi (top-200, en düşük q_diff)

| Karşılaştırma | EAC probe Jaccard | ESCC probe Jaccard |
|---------------|-------------------|---------------------|
| 10% vs 20% | 1.00 | 1.00 |
| 20% vs 30% | **0.311** | 1.00 |
| 10% vs 30% | 0.311 | 1.00 |

**EAC hasta üyeliği 20% vs 30%:** Jaccard = **0.714** (10/14 örtüşme).

### Panel C — Kilit lokuslar

- **ESCC SHISA3:** genome-wide q_diff **rank 1** @ 10%, 20%, 30% → metindeki "stable ESCC programme" iddiası **destekleniyor**.
- **EAC SIGLEC14, EPHB1, ZNF214:** rank 1–3 @ 10–20%, **rank düşüyor @ 30%** → "cut-off sensitive EAC programme" iddiası **destekleniyor**.

### Manuscript düzeltmesi gerekli

Sarı metinde **Jaccard 0.290** (EAC 20% vs 30%) yazıyor. Yeniden hesaplanan değer **0.311**. Gönderimden önce Methods/Results'ta **0.311** (veya "≈0.31") + **Supplementary Table S10** atfı kullanın.

Tam tablo metni: `submission_prep/supplementary_table_S10.md`

Ham veri:
- `result/fast_death_sensitivity_summary.tsv`
- `result/fast_death_sensitivity_key_loci.tsv`
- `result/fast_death_sensitivity_probes.tsv` (tam HM450; büyük dosya)

---

## 3. Data availability — manuscript'e yapıştırılacak metin

Dosya: **`submission_prep/data_availability_draft.txt`**

**Sizin yapmanız gereken (2 placeholder):**
1. `[INSERT PUBLIC REPOSITORY URL]` → GitHub public repo
2. `[INSERT ZENODO DOI]` → isteğe bağlı ama önerilir

Adım adım repo kurulumu: **`submission_prep/CODE_DEPOSIT_README.md`**

---

## 4. REMARK checklist — manuscript'e eklenecek

Dosya: **`submission_prep/remark_checklist.md`**

**Önerilen yerleşim:** Supplementary **Table S11** (veya Methods sonunda kısa paragraf + S11).

---

## 5. Response to Reviewer

Dosya: **`submission_prep/response_to_reviewer.md`**

- 18 Word yorumu (Sanne Johanna Maria, 26 May 2026)
- SR ilk gönderimde gerekmez; internal revizyon + dış hakem şablonu

---

## 6. Manuscript checklist (sizin)

- [ ] Data availability → `data_availability_draft.txt` + GitHub URL
- [ ] Supplementary Table S10 → `supplementary_table_S10.md`
- [ ] Supplementary Table S11 → REMARK checklist
- [ ] Jaccard **0.290 → 0.311**; isteğe bağlı patient Jaccard **0.714**
- [ ] Supplementary Materials listesine S10, S11

---

## 7. S10 yeniden üretim

```bash
cd "/Users/senol/Desktop/AML/Esophageal Carcinoma (TCGA, Firehose Legacy)"
python3 -m esca_escc_multiomic_pipeline.run_fast_death_sensitivity
```

---

## 8. Yapılmadı

- Cover letter (istek üzerine)
- Manuscript.docx düzenlemesi (siz)
- GitHub/Zenodo upload (sizin hesap)

Klasör: `submission_prep/`
