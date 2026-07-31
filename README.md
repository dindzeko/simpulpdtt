# Aplikasi Penarikan Kesimpulan Pemeriksaan Kepatuhan

Aplikasi Streamlit satu file (`app.py`) yang mereplikasi **Matriks Penyimpulan**
(Lampiran 6.1–6.4): pembobotan berjenjang Aspek → Sub-aspek/Kriteria → Temuan,
dengan rekalkulasi otomatis dan penarikan kesimpulan
(SESUAI / SESUAI DENGAN PENGECUALIAN / TIDAK SESUAI).

## Isi
- `app.py` — **satu file** berisi semuanya: engine perhitungan, generator Lampiran
  6.1–6.4, data contoh, dan antarmuka Streamlit. Cukup 1 file untuk dijalankan atau di-deploy.
- `requirements.txt`

## Menjalankan lokal
```bash
pip install -r requirements.txt
streamlit run app.py
```
Buka http://localhost:8501

## Deploy online (Streamlit Community Cloud — gratis)
1. Push folder ini ke repo GitHub (`app.py` di root).
2. Buka https://share.streamlit.io → **New app** → pilih repo & branch → main file `app.py`.
3. Deploy. URL publik otomatis dibuat; `requirements.txt` diinstal otomatis.

Alternatif server sendiri (mis. VPS/on-prem BPK):
```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true
```

## Logika perhitungan (per sub-aspek)
| Kolom | Rumus |
|---|---|
| Nilai Tertimbang (5) | `round(bobot_aspek × bobot_kriteria, 3)` |
| Skala Temuan (7) | rata-rata parameter aktif **atau** skala langsung (1–4) |
| Skor Temuan (8) | `round(nilai_tertimbang × skala, 2)` |
| Skor Penyimpangan (9) | `skor_temuan − nilai_tertimbang` |
| Batas Penyimpangan (10) | `round(batas × nilai_tertimbang, 3)` |
| Selisih (11) | `skor_penyimpangan − batas` |
| Kesimpulan (12) | `selisih ≥ 0 → MENYIMPANG` |

Kesimpulan keseluruhan dari Total Skor Temuan: `<1,3` SESUAI · `1,3–1,8`
SESUAI DENGAN PENGECUALIAN · `>1,8` TIDAK SESUAI (ambang dapat diubah di sidebar).

## Unduhan Lampiran
Tombol **"Unduh Lampiran 6.1–6.4 (form asli)"** menghasilkan satu workbook berisi
4 sheet dengan struktur & **formula hidup** persis aslinya:
- **6.1** Pembobotan Aspek (bobot + pertimbangan, Σ bobot = 1)
- **6.2** Pembobotan Sub-aspek (ref ke 6.1, subjumlah per aspek = 1)
- **6.3** Pembobotan Temuan (parameter → Jumlah → Rata-rata skala)
- **6.4** Kesimpulan (matriks penuh yang me-ref 6.1/6.2/6.3, total skor + kategori)

Auditor cukup mengisi data di aplikasi; seluruh lampiran ter-generate otomatis dan
akan menghitung ulang sendiri saat dibuka di Excel. Diverifikasi: 151 formula, 0 error.

## Penyempurnaan atas Excel
- Rata-rata skala dibagi **jumlah parameter aktif**, bukan selalu `/4`.
- Pembulatan memakai **ROUND ala Excel** (half-away-from-zero), bukan banker's
  rounding Python — hasil identik dengan workbook (Total = 1,66).
- Batas penyimpangan & ambang kesimpulan **configurable** di sidebar.
