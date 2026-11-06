---
title: Sistem Deteksi Kerusakan Gedung
emoji: 🏢
colorFrom: blue
colorTo: indigo
sdk: streamlit
sdk_version: 1.35.0
app_file: app.py
pinned: false
---

# SISTEM DETEKSI KERUSAKAN GEDUNG
### Aplikasi Skripsi Teknik Sipil Digital (Deteksi Kerusakan & Estimasi AHSP Jawa Tengah)

Aplikasi website komprehensif untuk inspeksi kondisi fisik bangunan gedung, identifikasi kerusakan otomatis menggunakan AI Agent Antigravity, serta perhitungan Rencana Anggaran Biaya (RAB) pemeliharaan mengacu pada standar resmi **AHSP Dinas PU Bina Marga & Cipta Karya Provinsi Jawa Tengah Tahun 2023** (berdasarkan Permen PUPR No. 1/PRT/M/2022 & Permen PUPR No. 24/PRT/M/2008).

---

## 🏛️ Spesifikasi Arsitektur Sistem

1. **Framework Utama**: Streamlit (Python) sebagai Frontend interaktif sekaligus Backend controller.
2. **Integrasi AI Agent**: Menggunakan library `requests` untuk melakukan HTTP POST JSON API Call ke AI Agent Antigravity dengan data citra terenkode Base64.
3. **Database Relasional**: SQLite3 (`riwayat_kerusakan`) untuk persistensi data hasil inspeksi dan penilaian kondisi gedung.
4. **Mesin Ekspor Excel (.xlsx)**: Menggunakan library `openpyxl` untuk menghasilkan lembar penilaian kondisi fisik dan fungsi (IKBG) sesuai format formulir standar teknik sipil:
   - Tabel Penilaian Fisik (Baik 79.5-100%, Cukup 49.5-79.49%, Rusak 0-49.49%)
   - Tabel Penilaian Fungsi (Baik, Cukup, Rusak)
   - Baris Sub Total Nilai Kondisi
   - Rumus Evaluasi Kinerja (µku): `Rutin: µku > 90; Preventif: µku = 70 - 90; Korektif: µku = 50 - 70; Rehabilitatif: µku < 50`
   - Tabel Bobot & Kinerja (Fisik 50%, Fungsi 50%)
   - Tabel Kotak Checkbox Status Nilai Kondisi ([X] RUTIN, PREVENTIF, KOREKTIF, REHABILITATIF)
   - Sheet Rekapitulasi Rincian RAB AHSP Jawa Tengah 2023
5. **Mesin Ekspor PDF**: ReportLab untuk memproduksi dokumen berita acara / laporan resmi pemeliharaan gedung lengkap dengan formulir IKBG, kop instansi, rekapitulasi RAB, dan lembar pengesahan tanda tangan tim penguji.
6. **UI/UX Design**: Desain modern bernuansa *Tailwind-CSS* dengan sistem badge tingkat keparahan, kartu metrik KPI, dan tabel terstruktur.

---

## 📁 Struktur Database SQLite (`pemeliharaan_gedung.db`)

Tabel: `riwayat_kerusakan`
- `id` (INTEGER PRIMARY KEY AUTOINCREMENT)
- `tanggal` (TEXT)
- `elemen_gedung` (TEXT)
- `jenis_kerusakan` (TEXT)
- `kode_ahsp` (TEXT)
- `volume` (REAL)
- `satuan` (TEXT)
- `harga_satuan` (INTEGER)
- `total_biaya` (INTEGER)
- `rekomendasi` (TEXT)

---

## 🚀 Cara Menjalankan Aplikasi

### Opsi A: Menggunakan Batch Script (Paling Praktis di Windows)
Cukup klik dua kali file:
- `run_app.bat`

### Opsi B: Melalui Terminal / Command Prompt
```bash
python -m streamlit run app.py
```

Aplikasi akan otomatis terbuka pada browser Anda di alamat: `http://localhost:8501`.

---

## 🤖 Menjalankan Server Mock AI Agent (Opsional)
Jika ingin mendemonstrasikan komunikasi jaringan live HTTP POST antara dua terminal terpisah:
1. Jalankan server mock AI di Terminal 1:
   ```bash
   python mock_ai_agent_server.py
   # atau klik ganda run_mock_ai_server.bat
   ```
2. Jalankan aplikasi Streamlit di Terminal 2:
   ```bash
   python -m streamlit run app.py
   ```
3. Endpoint `http://localhost:8000/api/analyze-damage` akan otomatis menerima HTTP POST dari aplikasi Streamlit.

*Catatan: Jika server mock tidak dijalankan, aplikasi tetap berjalan 100% aman berkat fitur **Auto-Fallback Standalone AI Engine** sehingga presentasi sidang skripsi dijamin tidak akan terganggu masalah jaringan.*

---

## 📷 Citra Sampel Uji Coba Lapangan
Tersedia folder `sample_images/` berisi citra uji coba teknik sipil:
- `retak_balok_beton.jpg` (Kerusakan Struktural Balok - AHSP Injeksi Epoxy)
- `retak_dinding_plesteran.jpg` (Kerusakan Dinding Pasangan - AHSP Mortar Polimer)
- `spalling_kolom_beton.jpg` (Spalling Selimut Beton & Karat Tulangan - AHSP Patching Mortar)
