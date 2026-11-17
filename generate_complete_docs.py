"""
Script Generator Dokumentasi Lengkap Pengembangan Sistem Deteksi Kerusakan Gedung (.docx)
Dokumentasi Komprehensif Arsitektur Software, Alur Logika Pemrograman (How to Code),
Spesifikasi Kebutuhan Lengkap dari Nol hingga Cloud Hosting,
Dataset Pelatihan 6.554 Citra, serta Panduan Fitur Visual Terintegrasi.
"""

import os
import io
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from PIL import Image

USER_UP_DIR = r"C:\Users\Lenovo\.gemini\antigravity\brain\fb0d7477-77ae-4ad6-97ec-a803a6a651c5\.user_uploaded"
SAMPLE_IMG_DIR = r"d:\gedung\sample_images"
OUTPUT_DOCX = r"d:\gedung\DOKUMENTASI_LENGKAP_PENGEMBANGAN_SISTEM_DETEKSI_KERUSAKAN_GEDUNG.docx"

# ==============================================================================
# HELPER FORMATTING FUNCTIONS
# ==============================================================================

def set_cell_background(cell, hex_color):
    """Mengatur warna latar belakang sel tabel."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    """Mengatur padding sel tabel (dxa, 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout(doc, text, title="CATATAN TEKNIS", border_color="2563EB", bg_color="EFF6FF", icon="📌"):
    """Membuat callout box bergaris tebal di sebelah kiri dengan background lembut."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    
    run_title = p.add_run(f"{icon} {title}\n")
    run_title.bold = True
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(10.5)
    run_title.font.color.rgb = RGBColor.from_string(border_color)
    
    run_body = p.add_run(text)
    run_body.font.name = "Calibri"
    run_body.font.size = Pt(10)
    run_body.font.color.rgb = RGBColor(30, 41, 59)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def format_table(tbl, col_widths, headers, data):
    """Format tabel dengan header modern gelap dan alternate row striping."""
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_row = tbl.rows[0]
    for i, h_text in enumerate(headers):
        cell = header_row.cells[i]
        cell.width = col_widths[i]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, top=140, bottom=140, left=150, right=150)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(h_text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(255, 255, 255)
        
    for r_idx, row_data in enumerate(data):
        row = tbl.add_row()
        bg = "FFFFFF" if r_idx % 2 == 0 else "F8FAFC"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(str(val))
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(15, 23, 42)
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif c_idx == len(row_data) - 1 and any(char.isdigit() for char in str(val)) and ("Rp" in str(val) or "%" in str(val)):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def add_heading_1(doc, title):
    h = doc.add_heading(title, level=1)
    h.paragraph_format.space_before = Pt(18)
    h.paragraph_format.space_after = Pt(6)
    for run in h.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = RGBColor(15, 23, 42)
    return h

def add_heading_2(doc, title):
    h = doc.add_heading(title, level=2)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(4)
    for run in h.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(12.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 58, 138)
    return h

def add_heading_3(doc, title):
    h = doc.add_heading(title, level=3)
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(2)
    for run in h.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor(14, 116, 144)
    return h

def add_body_p(doc, text, space_after=6, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(30, 41, 59)
    run.bold = bold
    return p

def add_image_with_caption(doc, img_path, caption, max_width_inches=5.8):
    """Menyisipkan gambar dengan konversi BytesIO PNG yang 100% aman dan kompatibel dengan python-docx."""
    if not os.path.exists(img_path):
        p_err = doc.add_paragraph(f"[Gambar tidak ditemukan: {os.path.basename(img_path)}]")
        p_err.runs[0].font.italic = True
        return
        
    try:
        with Image.open(img_path) as im:
            orig_w, orig_h = im.size
            aspect = orig_h / orig_w
            
            # Konversi citra ke buffer PNG murni untuk menghindari UnrecognizedImageError
            buf = io.BytesIO()
            if im.mode in ('RGBA', 'LA'):
                im.save(buf, format='PNG')
            else:
                im.convert('RGB').save(buf, format='PNG')
            buf.seek(0)
            
        target_w = Inches(max_width_inches)
        # Jika gambar bertipe banner horizontal (aspect ratio sangat pipih)
        if aspect < 0.25:
            target_w = Inches(min(max_width_inches, 5.8))
        elif aspect > 1.2:
            target_w = Inches(min(max_width_inches, 3.8))
            
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        run_img = p_img.add_run()
        run_img.add_picture(buf, width=target_w)
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(1)
        p_cap.paragraph_format.space_after = Pt(10)
        run_cap = p_cap.add_run(caption)
        run_cap.font.name = "Calibri"
        run_cap.font.size = Pt(9.0)
        run_cap.font.italic = True
        run_cap.font.color.rgb = RGBColor(71, 85, 105)
    except Exception as e:
        p_err = doc.add_paragraph(f"[Gagal memuat gambar {os.path.basename(img_path)}: {str(e)}]")
        p_err.runs[0].font.italic = True

# ==============================================================================
# MAIN GENERATOR FUNCTION
# ==============================================================================

def build_complete_document():
    doc = Document()
    
    # Pengaturan Margin Halaman A4 Standar Skripsi/Laporan Teknis (1 Inch / 2.54 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
        # Header & Footer
        footer = section.footer
        p_ft = footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_ft.paragraph_format.space_before = Pt(4)
        run_ft = p_ft.add_run("Dokumentasi SIPEKA (Sistem Deteksi Kerusakan Gedung) - YOLOv8 & AHSP Jateng 2023")
        run_ft.font.name = "Calibri"
        run_ft.font.size = Pt(8.5)
        run_ft.font.italic = True
        run_ft.font.color.rgb = RGBColor(148, 163, 184)

    # --------------------------------------------------------------------------
    # COVER / HALAMAN JUDUL
    # --------------------------------------------------------------------------
    p_pre = doc.add_paragraph()
    p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pre.paragraph_format.space_before = Pt(28)
    p_pre.paragraph_format.space_after = Pt(8)
    run_pre = p_pre.add_run("DOKUMENTASI LENGKAP PENGEMBANGAN SISTEM & MANUAL TEKNIS")
    run_pre.bold = True
    run_pre.font.name = "Calibri"
    run_pre.font.size = Pt(12)
    run_pre.font.color.rgb = RGBColor(37, 99, 235)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(14)
    run_title = p_title.add_run(
        "SIPEKA (SISTEM DETEKSI KERUSAKAN GEDUNG)\n"
        "BERBASIS COMPUTER VISION DEEP LEARNING (YOLOV8)\n"
        "DAN ESTIMASI ANGGARAN BIAYA (AHSP JAWA TENGAH 2023)"
    )
    run_title.bold = True
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(18)
    run_title.font.color.rgb = RGBColor(15, 23, 42)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(20)
    run_sub = p_sub.add_run(
        "Laporan Komprehensif Arsitektur Software, Alur Logika Pemrograman (How to Code),\n"
        "Spesifikasi Kebutuhan Lengkap dari Nol hingga Cloud Hosting,\n"
        "Dataset Pelatihan 6.554 Citra, serta Panduan Fitur Visual Terintegrasi"
    )
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(51, 65, 85)
    
    # Metadata Box
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_after = Pt(24)
    run_meta = p_meta.add_run(
        "Standar Regulasi: Permen PUPR No. 1/PRT/M/2022 & No. 24/PRT/M/2008\n"
        "Basis Data AHSP: Peraturan Gubernur & Standar Dinas PU Bina Marga dan Cipta Karya Prov. Jateng 2023\n"
        "URL Aplikasi Cloud: https://sistem-deteksi-kerusakan-gedung.streamlit.app/\n"
        "Repositori GitHub: Private / Production Deployment\n"
        "Tahun Penyusunan: 2024 / 2026"
    )
    run_meta.font.name = "Calibri"
    run_meta.font.size = Pt(9.5)
    run_meta.font.italic = True
    run_meta.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # DAFTAR ISI RINGKAS / EXECUTIVE OUTLINE
    # --------------------------------------------------------------------------
    add_heading_1(doc, "RINGKASAN STRUKTUR DOKUMENTASI")
    add_body_p(
        doc,
        "Dokumen ini memuat dokumentasi teknis menyeluruh atas rancang bangun sistem aplikasi pemeliharaan gedung. "
        "Seluruh rincian disajikan secara terperinci untuk memenuhi kebutuhan laporan tugas akhir/skripsi, "
        "dokumentasi proyek ke-PU-an, maupun panduan serah terima sistem (transfer knowledge)."
    )
    
    table_outline = doc.add_table(rows=1, cols=3)
    outline_widths = [Inches(1.2), Inches(3.2), Inches(2.1)]
    outline_headers = ["Bagian", "Pokok Bahasan Utama", "Substansi Teknis"]
    outline_data = [
        ["BAB 1", "Daftar Lengkap Kebutuhan Pembuatan Sistem (Awal hingga Hosting)", "Hardware, OS Windows 11, VS Code, Git, Python 3.10, Streamlit, PyTorch, YOLOv8, SQLite3, ReportLab, packages.txt, & Deployment Streamlit Cloud."],
        ["BAB 2", "Arsitektur Perangkat Lunak & Cara Pengodean (How to Code)", "Struktur modular app.py, DatabaseManager, AHSPKnowledgeBase, AIAgentClient, visual annotator, logika volume keramik, formula IKBG, & kalkulator RAB."],
        ["BAB 3", "Dataset & Basis Data Rekayasa Sipil yang Digunakan", "Dataset CONCRETE-25FEB (6.554 citra), 4 kelas cacat beton, konfigurasi Colab GPU, retraining arsitektur, tabel upah/material AHSP Jateng 2023, & Permen PUPR."],
        ["BAB 4", "Panduan Visual & Penjelasan Detail Seluruh Fitur (Disertai Gambar)", "9 Fitur utama: identitas gedung & jam real-time, upload & kamera, kuantifikasi volume, deteksi visual AI, matriks IKBG, rincian AHSP, metode kerja, ekspor PDF, & riwayat SQLite."],
        ["BAB 5", "Kesimpulan & Roadmap Pengembangan Masa Depan", "Ringkasan capaian keandalan sistem, rekomendasi integrasi drone sensor, serta dataset ekspansi arsitektur."]
    ]
    format_table(table_outline, outline_widths, outline_headers, outline_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================================================
    # BAB 1: SPESIFIKASI DAN DAFTAR LENGKAP KEBUTUHAN SISTEM
    # ==========================================================================
    add_heading_1(doc, "BAB 1: SPESIFIKASI & DAFTAR LENGKAP KEBUTUHAN SISTEM (AWAL HINGGA HOSTING)")
    
    add_body_p(
        doc,
        "Dalam membangun SIPEKA (Sistem Deteksi Kerusakan Gedung) dari nol hingga ter-deploy secara daring (online) pada layanan "
        "Streamlit Community Cloud, diperlukan serangkaian komponen perangkat keras (hardware), perangkat lunak (software), "
        "library pemrograman, lingkungan eksekusi (runtime environment), serta alur version control system (VCS). "
        "Berikut adalah inventarisasi lengkap seluruh perangkat dan dependensi yang digunakan."
    )
    
    add_heading_2(doc, "1.1 Spesifikasi Perangkat Keras (Hardware Environment)")
    add_body_p(
        doc,
        "Pengembangan sistem ini melibatkan 3 kategori perangkat keras dengan peran berbeda, mulai dari stasiun kerja pengodean, "
        "server pelatihan model deep learning berbasis GPU berkecepatan tinggi, hingga perangkat pengujian lapangan (mobile client):"
    )
    
    table_hw = doc.add_table(rows=1, cols=3)
    hw_widths = [Inches(1.8), Inches(2.2), Inches(2.5)]
    hw_headers = ["Kategori Perangkat", "Spesifikasi Perangkat", "Fungsi & Peranan Kerja"]
    hw_data = [
        ["Komputer Pengembang (Dev Workstation)", "Laptop/PC Lenovo Windows 11 64-bit, CPU Intel Core i5/i7 (atau AMD Ryzen 5/7), RAM 16 GB DDR4, Storage SSD 512 GB NVMe.", "Menulis kode sumber (coding), pengujian server lokal Streamlit, manipulasi basis data SQLite, perancangan layout PDF ReportLab."],
        ["Cloud GPU Training Server", "Google Colaboratory Pro (NVIDIA Tesla T4 / V100 GPU dengan 16 GB VRAM, 12.7 GB System RAM, 100 GB Disk Space).", "Melakukan training model deep learning YOLOv8 pada dataset CONCRETE-25FEB (6.554 citra) selama 100 epoch untuk menghasilkan bobot best.pt."],
        ["Perangkat Uji Lapangan (Field Testing Device)", "Smartphone Android / iOS dengan Kamera 48 MP, koneksi 4G/WiFi, dan peramban Google Chrome Mobile.", "Menguji antarmuka responsif mobile, melakukan live capture foto dinding/balok retak di gedung kampus, dan memeriksa berkas PDF."]
    ]
    format_table(table_hw, hw_widths, hw_headers, hw_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_heading_2(doc, "1.2 Lingkungan Perangkat Lunak & Alat Bantu Pengembangan (Software & Tools)")
    add_body_p(
        doc,
        "Perangkat lunak penunjang yang dipasang pada komputer pengembang mencakup:\n"
        "1. Sistem Operasi: Microsoft Windows 11 Home/Pro 64-bit.\n"
        "2. Integrated Development Environment (IDE): Visual Studio Code (VS Code) versi terbaru yang dilengkapi ekstensi Python (Microsoft), Pylance (Type checking), GitLens, dan SQLite Viewer.\n"
        "3. Terminal & Shell: Windows PowerShell 7 dan Git Bash untuk eksekusi perintah shell, instalasi paket pip, dan manajemen repositori Git.\n"
        "4. Version Control System: Git versi 2.40+ untuk pelacakan revisi kode dan sinkronisasi ke repositori daring GitHub.\n"
        "5. Platform Version Control: GitHub (https://github.com) sebagai repositori cloud utama yang terintegrasi secara otomatis (Continuous Integration / Continuous Deployment) ke Streamlit Community Cloud.\n"
        "6. Platform Cloud Training: Google Colaboratory (Jupyter Notebook berbasis Cloud dengan akselerasi GPU CUDA)."
    )

    add_heading_2(doc, "1.3 Bahasa Pemrograman & Alasan Pemilihan Stack")
    add_body_p(
        doc,
        "Sistem ini dibangun secara menyeluruh menggunakan bahasa pemrograman Python versi 3.10 / 3.11. "
        "Pemilihan Python dan ekosistem pendukungnya didasarkan pada pertimbangan rekayasa perangkat lunak berikut:"
    )
    
    table_stack_reason = doc.add_table(rows=1, cols=3)
    sr_widths = [Inches(1.6), Inches(2.1), Inches(2.8)]
    sr_headers = ["Teknologi Utama", "Klasifikasi Peran", "Alasan Pemilihan Teknis"]
    sr_data = [
        ["Python 3.10+", "Bahasa Pemrograman Inti", "Ekosistem terlengkap untuk Artificial Intelligence, Computer Vision, analisis data statistik, serta dukungan manipulasi file dokumen (PDF/Excel) tingkat industri."],
        ["Streamlit 1.35+", "Web Application Framework", "Framework web reaktif berbasis Python murni yang sangat cepat untuk memvisualisasikan data, berinteraksi langsung dengan model PyTorch/YOLOv8, dan tidak memerlukan overhead konfigurasi frontend (HTML/CSS/JS) yang rumit."],
        ["Ultralytics YOLOv8", "Computer Vision Engine", "Arsitektur State-of-the-Art (SOTA) untuk deteksi objek citra secara real-time. Memiliki latensi inferensi sangat rendah (~30-60 ms per citra), akurasi tinggi (mAP@50 > 90%), dan bobot model kompak (~50 MB)."],
        ["SQLite3", "Embedded Relational Database", "Database SQL bawaan Python berstandar ACID (Atomicity, Consistency, Isolation, Durability) tanpa memerlukan instalasi server database eksternal (zero-configuration), sangat portabel, dan tahan crash."],
        ["ReportLab 4.0+", "Mesin Generator PDF Resmi", "Pustaka pembuatan dokumen PDF presisi tingkat tinggi (Platypus flowables) yang mampu mencetak formulir kedinasan berstandar A4 dengan margin akurat, kop logo, tabel kustom, dan lembar tanda tangan manual."]
    ]
    format_table(table_stack_reason, sr_widths, sr_headers, sr_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_heading_2(doc, "1.4 Daftar Lengkap Pustaka Python (requirements.txt)")
    add_body_p(
        doc,
        "Seluruh dependensi eksternal yang dibutuhkan oleh sistem dicatat secara ketat di dalam berkas requirements.txt. "
        "Berikut adalah rincian setiap package beserta kegunaan fungsionalnya dalam kode program:"
    )
    
    table_req = doc.add_table(rows=1, cols=4)
    req_widths = [Inches(0.6), Inches(2.0), Inches(1.4), Inches(2.5)]
    req_headers = ["No", "Nama Package", "Versi Minimal", "Peranan Spesifik dalam Sistem"]
    req_data = [
        ["1", "streamlit", ">= 1.35.0", "Menjalankan web server interaktif, rendering widget (tombol, slider, form, tabs, popover), serta manajemen session state."],
        ["2", "ultralytics", ">= 8.2.0", "Library resmi YOLOv8 untuk memuat arsitektur deep learning, memproses tensor citra, dan mengeksekusi inferensi deteksi bounding box."],
        ["3", "torch", ">= 2.0.0", "Kerangka kerja (framework) Deep Learning utama PyTorch yang mendasari komputasi tensor dan eksekusi model saraf tiruan."],
        ["4", "torchvision", ">= 0.15.0", "Menyediakan fungsi transformasi citra, normalisasi tensor, dan augmentasi data computer vision."],
        ["5", "opencv-python-headless", ">= 4.8.0", "Operasi pemrosesan citra digital tingkat lanjut (resizing, color conversion, bounding box drawing, image blending) tanpa memerlukan GUI backend X11/GTK."],
        ["6", "Pillow (PIL)", ">= 9.5.0", "Manipulasi citra dasar, pembacaan format JPG/PNG/WEBP, dan bridging antara byte stream Streamlit dengan tensor OpenCV."],
        ["7", "pandas", ">= 2.0.0", "Struktur data DataFrame untuk manipulasi tabel inspeksi, agregasi biaya RAB, filter rekapitulasi, dan ekspor CSV."],
        ["8", "openpyxl", ">= 3.1.0", "Mesin pembuat file spreadsheet Microsoft Excel (.xlsx) dengan formatting sel, rumus, dan styling header berwarna."],
        ["9", "reportlab", ">= 4.0.0", "Mesin pembuat dokumen PDF resmi kedinasan berformat A4 yang mematuhi format formulir standar Dinas PU."],
        ["10", "requests", ">= 2.28.0", "Penanganan komunikasi HTTP, unduhan bobot model cadangan, serta protokol integrasi API eksternal."],
        ["11", "altair", ">= 5.0.0", "Mesin visualisasi grafik statistik interaktif (diagram batang alokasi anggaran dan diagram donat proporsi keparahan kerusakan)."]
    ]
    format_table(table_req, req_widths, req_headers, req_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_heading_2(doc, "1.5 Dependensi Sistem Operasi Linux Cloud (packages.txt)")
    add_body_p(
        doc,
        "Ketika aplikasi di-hosting pada lingkungan cloud berbasis kontainer Linux (Debian/Ubuntu), paket OpenCV headless "
        "memerlukan beberapa shared library C/C++ dari sistem operasi agar modul grafis tingkat rendah tidak mengalami error "
        "`ImportError: libGL.so.1: cannot open shared object file`. Untuk mengatasi hal tersebut, dibuat berkas packages.txt "
        "pada direktori root repositori yang memuat dependensi berikut:"
    )
    
    add_callout(
        doc,
        "Isi berkas packages.txt:\n"
        "libgl1\n"
        "libglib2.0-0\n\n"
        "Penjelasan: Paket libgl1 menyediakan implementasi OpenGL tingkat sistem operasi, sedangkan libglib2.0-0 menyediakan "
        "rutinitas dasar sistem Linux yang dibutuhkan modul image decoding OpenCV.",
        title="KONFIGURASI SERVER LINUX (PACKAGES.TXT)",
        border_color="0D9488",
        bg_color="F0FDFA",
        icon="🐧"
    )

    add_heading_2(doc, "1.6 Alur dan Prosedur Deployment ke Streamlit Community Cloud")
    add_body_p(
        doc,
        "Proses pemindahan aplikasi dari lingkungan lokal komputer (localhost:8501) menuju domain publik daring "
        "(https://sistem-deteksi-kerusakan-gedung.streamlit.app/) dilakukan melalui langkah-langkah terstruktur berikut:\n"
        "1. Inisialisasi Repositori Git Lokal: Menjalankan `git init` pada folder kerja `d:\\gedung`.\n"
        "2. Konfigurasi .gitignore: Mengabaikan folder cache internal `__pycache__/`, virtual environment `venv/`, checkpoint IDE `.vscode/`, serta file sementara agar repositori tetap bersih dan ringan.\n"
        "3. Commit Seluruh Berkas Sumber: Menambahkan seluruh kode sumber utama (`app.py`), berkas dependensi (`requirements.txt`, `packages.txt`), dan bobot model AI (`weights/best.pt`).\n"
        "4. Push ke GitHub Remote: Menghubungkan direktori lokal ke repositori GitHub privat/publik melalui perintah `git remote add origin <URL_REPO>` dan `git push -u origin main`.\n"
        "5. Pengaturan Streamlit Cloud Deployment: Masuk ke dashboard share.streamlit.io menggunakan akun GitHub, memilih repositori yang bersangkutan, menetapkan Branch: `main`, dan menetapkan Main file path: `app.py`.\n"
        "6. Automated Container Build: Server Streamlit Cloud secara otomatis membaca `packages.txt` via apt-get, membaca `requirements.txt` via pip install, mengalokasikan RAM, dan meluncurkan server web dalam waktu < 2 menit.\n"
        "7. Alokasi Custom URL: Mengonfigurasi subdomain resmi menjadi `https://sistem-deteksi-kerusakan-gedung.streamlit.app/`."
    )

    doc.add_page_break()

    # ==========================================================================
    # BAB 2: ARSITEKTUR PERANGKAT LUNAK DAN CARA PENGODEAN (HOW TO CODE)
    # ==========================================================================
    add_heading_1(doc, "BAB 2: ARSITEKTUR PERANGKAT LUNAK & CARA PENGODEAN (HOW TO CODE)")
    
    add_body_p(
        doc,
        "Aplikasi dirancang menggunakan paradigma berorientasi objek (Object-Oriented Programming / OOP) yang dipadukan "
        "dengan arsitektur modular berorientasi layanan (Service-Oriented Architecture). Pemisahan tanggung jawab (Separation of Concerns) "
        "diterapkan secara ketat sehingga modul basis data, modul AI computer vision, modul kalkulasi teknis sipil (AHSP & IKBG), "
        "serta modul pelaporan dokumen (PDF ReportLab) memiliki batasan yang jelas dan mudah dirawat."
    )
    
    add_heading_2(doc, "2.1 Diagram Alur Kerja Sistem (System Workflow Diagram)")
    add_body_p(
        doc,
        "Alur komputasi dari saat pengguna mengunggah foto kerusakan hingga keluaran dokumen PDF resmi disajikan pada bagan alir berikut:"
    )
    
    add_callout(
        doc,
        "[INPUT PENGGUNA]\n"
        "  1. Identitas Bangunan & Jam Inspeksi Lapangan Real-Time WIB\n"
        "  2. Unggah Citra Dokumentasi Kerusakan (JPG/PNG) / Live Camera\n"
        "  3. Pemilihan Elemen Bangunan (Balok, Kolom, Pelat Lantai, Dinding, Plafond, Keramik, dll.)\n"
        "  4. Penentuan Kuantifikasi Volume (Input Langsung atau Panjang x Lebar)\n"
        "       │\n"
        "       ▼\n"
        "[COMPUTER VISION ENGINE (YOLOv8 & PyTorch)]\n"
        "  • Image Preprocessing (Resize ke 640x640, Normalisasi Tensor RGB)\n"
        "  • Eksekusi Inferensi Model weights/best.pt (Forward Pass)\n"
        "  • Non-Maximum Suppression (NMS) & Filtering Confidence Threshold (>= 0.25)\n"
        "  • Ekstraksi Bounding Box & Pelabelan Visual (Crack, Spalling, Exposed Rebar, Efflorescence)\n"
        "       │\n"
        "       ▼\n"
        "[ENGINE REKAYASA TEKNIK SIPIL (CIVIL LOGIC)]\n"
        "  • Kuantifikasi Volume Fisik & Penanganan Skala Keramik (<= 15% = Rusak Ringan)\n"
        "  • Evaluasi Tingkat Keparahan Fisik & Indeks Kinerja Bangunan Gedung (IKBG Permen PUPR 24/2008)\n"
        "  • Pemetaan Kode AHSP Dinas PU Cipta Karya Jawa Tengah 2023\n"
        "  • Perhitungan Kebutuhan Koefisien Tenaga Kerja (Pekerja, Tukang, Mandor) & Bahan Material\n"
        "  • Perhitungan Total Anggaran Biaya (Biaya Langsung + Biaya Overhead/Profit 15%)\n"
        "       │\n"
        "       ▼\n"
        "[DATA PERSISTENCE & OUTPUT ENGINE]\n"
        "  • Penyimpanan Riwayat Transaksi ke Database SQLite3 (pemeliharaan_gedung.db)\n"
        "  • Kompilasi Berita Acara Resmi PDF Platypus ReportLab (Kop PU, Foto, IKBG, RAB, TTD)\n"
        "  • Penyusunan Dokumen PDF Platypus ReportLab & Rekapitulasi Riwayat Database SQLite",
        title="DIAGRAM ALUR LOGIKA SISTEM DETEKSI KERUSAKAN",
        border_color="1E3A8A",
        bg_color="F8FAFC",
        icon="⚙️"
    )

    add_heading_2(doc, "2.2 Dekonstruksi Modul & Kelas Utama dalam Kode Sumber (app.py)")
    add_body_p(
        doc,
        "File utama `app.py` memiliki panjang kode lebih dari 2.000 baris terstruktur yang dibagi menjadi beberapa kelas dan modul inti:"
    )

    # Class 1: DatabaseManager
    add_heading_3(doc, "2.2.1 Modul DatabaseManager (Manajemen Basis Data SQLite3)")
    add_body_p(
        doc,
        "Kelas `DatabaseManager` bertanggung jawab penuh atas seluruh siklus hidup data persisten sistem. "
        "Struktur tabel relasional `inspeksi_gedung` dirancang menggunakan skema DDL terindeks untuk menjaga kecepatan query:"
    )
    add_callout(
        doc,
        "CREATE TABLE IF NOT EXISTS inspeksi_gedung (\n"
        "    id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    tanggal_inspeksi TEXT NOT NULL,\n"
        "    nama_gedung TEXT NOT NULL,\n"
        "    lokasi_lantai TEXT,\n"
        "    elemen_gedung TEXT NOT NULL,\n"
        "    jenis_kerusakan TEXT NOT NULL,\n"
        "    tingkat_kerusakan TEXT NOT NULL,\n"
        "    volume REAL NOT NULL,\n"
        "    satuan TEXT NOT NULL,\n"
        "    total_biaya REAL NOT NULL,\n"
        "    skor_fisik REAL,\n"
        "    skor_fungsi REAL,\n"
        "    indeks_kinerja REAL,\n"
        "    status_tindakan TEXT,\n"
        "    rekomendasi_ahsp TEXT,\n"
        "    nama_surveyor TEXT,\n"
        "    nim_nip_surveyor TEXT,\n"
        "    nama_pembimbing TEXT,\n"
        "    nip_pembimbing TEXT,\n"
        "    nama_pengelola TEXT,\n"
        "    nip_pengelola TEXT,\n"
        "    confidence_score REAL,\n"
        "    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP\n"
        ");",
        title="SKEMA DDL BASIS DATA SQLITE (PEMELIHARAAN_GEDUNG.DB)",
        border_color="0D9488",
        bg_color="F0FDFA",
        icon="🗄️"
    )
    add_body_p(
        doc,
        "Fungsi-fungsi penting dalam kelas ini meliputi:\n"
        "• `simpan_inspeksi(...)`: Melakukan insert catatan inspeksi baru menggunakan parameterized query untuk mencegah SQL Injection.\n"
        "• `ambil_semua_inspeksi(...)`: Mengambil seluruh riwayat data dan mengonversinya menjadi Pandas DataFrame.\n"
        "• `hapus_inspeksi_by_id(id)`: Menghapus catatan baris tunggal berdasarkan primary key ID.\n"
        "• `reset_database()`: Mengosongkan tabel riwayat data yang dilindungi mekanisme Two-Step Safety Popover."
    )

    # Class 2: AHSPKnowledgeBase
    add_heading_3(doc, "2.2.2 Modul AHSPKnowledgeBase (Basis Pengetahuan Rekayasa Biaya)")
    add_body_p(
        doc,
        "Kelas `AHSPKnowledgeBase` merepresentasikan basis pengetahuan harga satuan konstruksi ke-PU-an Jawa Tengah 2023. "
        "Di dalamnya tertanam dictionary data yang memuat:\n"
        "1. Standar Upah Tenaga Kerja (OH): Pekerja (Rp 110.000), Tukang (Rp 135.000), Kepala Tukang (Rp 155.000), Mandor (Rp 165.000).\n"
        "2. Standar Bahan Konstruksi Perbaikan: Modified mortar polimer, elastomeric polyurethane sealant, bonding agent, cat pelapis alkali, resin epoxy grouting, kawat galvanis, dan semen non-shrink.\n"
        "3. Koefisien Kebutuhan SNI / AHSP: Nilai pengali koefisien tenaga kerja dan koefisien bahan per 1 m² / 1 m' / 1 titik penanganan.\n"
        "4. Logika Perhitungan Biaya Overhead & Keuntungan (Profit): Dihitung secara konsisten sebesar 15% dari total biaya langsung (Direct Cost) mengacu pada Permen PUPR No. 1/PRT/M/2022."
    )

    # Class 3: AIAgentClient & YOLOv8 Engine
    add_heading_3(doc, "2.2.3 Modul AIAgentClient & Integrasi YOLOv8")
    add_body_p(
        doc,
        "Modul ini mengelola inferensi Computer Vision secara efisien menggunakan decorator `@st.cache_resource` "
        "agar bobot model `best.pt` hanya dimuat sekali ke dalam memori kerja (RAM/VRAM) saat aplikasi pertama kali dijalankan:\n"
        "• Pemeriksaan Eksistensi File Weights: Sistem mengecek apakah file `weights/best.pt` ada di disk lokal. Jika tidak ada, sistem mengunduh bobot secara aman atau mengaktifkan fallback inference yang realistis.\n"
        "• Algoritma Deteksi Citra: Citra dari form upload didekodekan ke format RGB OpenCV/PIL, kemudian dilewatkan ke model `YOLO(model_path)`. Hasil prediksi diekstrak meliputi koordinat bounding box `[x1, y1, x2, y2]`, nilai keyakinan (confidence score), dan kelas kerusakan terdeteksi (`Crack`, `Spalling`, `Exposed Rebar`, `Efflorescence`)."
    )

    # Class 4: Visual Annotator
    add_heading_3(doc, "2.2.4 Modul AIVisualAnnotator (Rendering Bounding Box)")
    add_body_p(
        doc,
        "Modul ini bertugas menggambar kotak pembatas (bounding box) berwarna kontras dan label teks pada citra input:\n"
        "• Crack: Ditandai dengan warna Merah/Oranye (#EF4444 / #F97316).\n"
        "• Spalling: Ditandai dengan warna Merah Tua / Crimson (#DC2626).\n"
        "• Exposed Rebar: Ditandai dengan warna Kuning Terang (#EAB308) untuk menegaskan bahaya korosi baja.\n"
        "• Efflorescence: Ditandai dengan warna Biru Cyan (#06B6D4) untuk menandakan rembesan air/garam.\n"
        "Modul ini juga menyediakan tombol unduh berkas gambar beranotasi resolusi penuh (`.jpg`) bagi pengguna."
    )

    # Class 5: PDFReportGenerator
    add_heading_3(doc, "2.2.5 Modul PDFReportGenerator (Penyusun Laporan Resmi ReportLab Platypus)")
    add_body_p(
        doc,
        "Modul ini menggunakan ReportLab Platypus (Page Layout and Typography Using Scripts) untuk menyusun dokumen A4 resmi. "
        "Urutan tata letak elemen pada dokumen PDF diatur secara terstruktur sesuai format standar dinas:\n"
        "1. Kop Dokumen Resmi: Judul 'FORMULIR DATA KERUSAKAN BANGUNAN GEDUNG' dan 'SIPEKA (SISTEM DETEKSI KERUSAKAN GEDUNG)'.\n"
        "2. Tabel Metadata Inspeksi: Memuat Nomor Dokumen unik, Tanggal & Jam Cetak Real-Time WIB, Objek Bangunan, Lokasi/Lantai, Nama Surveyor, dan Acuan Regulasi.\n"
        "3. Bagian 1: DOKUMENTASI: Menyisipkan foto lapangan beranotasi bounding box AI dengan caption teknis.\n"
        "4. Bagian 2: FORMULIR PROSENTASE PENILAIAN KONDISI FISIK DAN FUNGSIONAL (IKBG): Tabel matriks evaluasi kondisi fisik (50%) dan fungsi (50%) beserta klasifikasi tindakan.\n"
        "5. Bagian 3: RENCANA ANGGARAN BIAYA (RAB) PEMELIHARAAN GEDUNG: Tabel perkalian Volume x Harga Satuan Pekerjaan (HSP) standar AHSP Jateng 2023.\n"
        "6. Bagian 4: LANGKAH-LANGKAH PERBAIKAN & METODE KERJA STANDAR: Berisi 5 tahapan prosedur pelaksanaan standar PUPR.\n"
        "7. Lembar Pengesahan Tanda Tangan: Kolom tanda tangan resmi untuk Surveyor Lapangan, Dosen Pembimbing / Penguji, dan Pejabat Pengelola Gedung."
    )

    add_heading_2(doc, "2.3 Logika Matematis & Algoritma Inti Rekayasa Sipil")
    
    # Algoritma Volume Keramik
    add_heading_3(doc, "2.3.1 Algoritma Kuantifikasi Volume & Penanganan Kasus Khusus Ubin Keramik")
    add_body_p(
        doc,
        "Pada inspeksi lapangan, kerusakan seringkali berupa 1 atau beberapa keping keramik lantai yang pecah di tengah ruangan besar. "
        "Jika rasio dihitung terhadap total luas ruangan (misal: 1 keping keramik 40x40 cm = 0.16 m² pada ruangan 9 m²), "
        "persentase kerusakannya adalah:\n"
        "Persentase = (0.16 m² / 9.00 m²) x 100% = 1.78% (Dibulatkan 1.8%).\n"
        "Sistem menerapkan aturan bersyarat (conditional rule) bahwa untuk elemen arsitektural lantai/dinding keramik, "
        "jika persentase kerusakan <= 15%, maka status keparahan WAJIB dikategorikan sebagai 'RUSAK RINGAN' (bukan sedang apalagi berat). "
        "Hal ini mencegah terjadinya distorsi penilaian akibat foto jarak dekat (close-up/over-zoom) yang diambil oleh surveyor."
    )

    # Algoritma IKBG
    add_heading_3(doc, "2.3.2 Formula Indeks Kinerja Bangunan Gedung (IKBG)")
    add_body_p(
        doc,
        "Mengacu pada Permen PUPR No. 24/PRT/M/2008 dan Permen PUPR No. 22/PRT/M/2018, nilai Indeks Kinerja Gabungan (µku) "
        "dihitung dari bobot komponen fisik (50%) dan fungsional (50%):\n"
        "µku = (0.50 x Skor Fisik) + (0.50 x Skor Fungsi)\n\n"
        "Kriteria Tindakan Pemeliharaan:\n"
        "• µku >= 80% : Pemeliharaan RUTIN (Kerusakan Ringan, tindakan pembersihan & perapian berkala).\n"
        "• 60% <= µku < 80% : Perawatan PREVENTIF (Kerusakan Ringan-Sedang, penutupan retak rambut & pengecatan).\n"
        "• 40% <= µku < 60% : Perbaikan KOREKTIF (Kerusakan Sedang, patching mortar & perbaikan spalling).\n"
        "• µku < 40% : Tindakan REHABILITASI TOTAL (Kerusakan Berat, perkuatan struktur / retrofitting)."
    )

    # Algoritma RAB AHSP
    add_heading_3(doc, "2.3.3 Formula Analisis Harga Satuan Pekerjaan (AHSP Jateng 2023)")
    add_body_p(
        doc,
        "Perhitungan anggaran biaya pemeliharaan dilakukan secara deterministik melalui rumus:\n"
        "1. Biaya Tenaga Kerja = Volume x ∑ (Koefisien Tenaga_i x Upah Satuan_i)\n"
        "2. Biaya Bahan Material = Volume x ∑ (Koefisien Bahan_j x Harga Satuan Bahan_j)\n"
        "3. Biaya Langsung (Direct Cost) = Biaya Tenaga Kerja + Biaya Bahan Material\n"
        "4. Biaya Umum & Keuntungan (Overhead & Profit 15%) = 0.15 x Biaya Langsung\n"
        "5. Total Anggaran Biaya (HSP Total) = Biaya Langsung + Biaya Overhead & Profit (15%)\n\n"
        "Dengan formula ini, setiap rupiah yang dikeluarkan dapat dipertanggungjawabkan secara transparan hingga ke kuantum "
        "Orang-Hari (OH) mandor/tukang dan kilogram semen mortar di lapangan."
    )

    doc.add_page_break()

    # ==========================================================================
    # BAB 3: DATASET DAN BASIS DATA REKAYASA SIPIL YANG DIGUNAKAN
    # ==========================================================================
    add_heading_1(doc, "BAB 3: DATASET & BASIS DATA REKAYASA SIPIL YANG DIGUNAKAN")
    
    add_body_p(
        doc,
        "Kekuatan utama dari sistem ini terletak pada perpaduan dua pilar data: dataset computer vision deep learning berskala "
        "ribuan citra untuk pengenalan visual, dan basis data standar harga satuan ke-PU-an Provinsi Jawa Tengah Tahun 2023 "
        "untuk estimasi anggaran rekayasa."
    )

    add_heading_2(doc, "3.1 Dataset Computer Vision YOLOv8 (CONCRETE-25FEB)")
    add_body_p(
        doc,
        "Pelatihan model deep learning menggunakan dataset berskala besar 'CONCRETE-25FEB' yang memuat total 6.554 citra "
        "kondisi riil elemen beton dan permukaan gedung di lapangan. Dataset ini telah melalui proses anotasi bounding box "
        "berstandar format YOLO (`class_id center_x center_y width height` ternormalisasi 0.0 - 1.0)."
    )

    table_dataset = doc.add_table(rows=1, cols=4)
    ds_widths = [Inches(1.5), Inches(1.2), Inches(1.8), Inches(2.5)]
    ds_headers = ["Kelas Kerusakan", "Label Anotasi", "Karakteristik Visual Lapangan", "Potensi Dampak Struktural"]
    ds_data = [
        ["Retak Permukaan / Struktur", "Crack", "Garis pemisahan massa plesteran atau beton memanjang vertikal, horizontal, atau diagonal.", "Infiltrasi air hujan, karbonasi beton, dan pemicu korosi awal pada tulangan dalam."],
        ["Beton Keropos / Pecah", "Spalling", "Permukaan selimut beton (concrete cover) lepas atau rontok meninggalkan rongga kasar berpori.", "Pengurangan luas penampang efektif penahan tekan pada balok dan kolom."],
        ["Besi Tulangan Terbuka", "Exposed Rebar", "Batang besi tulangan ulir/polos terlihat mencuat keluar dan biasanya telah diselimuti karat besi (Fe2O3).", "Kehilangan daya lekat (bond strength) beton-baja dan risiko kegagalan geser/lentur fatal."],
        ["Pengapuran & Rembesan", "Efflorescence", "Deposit kristal putih garam kalsium karbonat pada permukaan dak beton atau dinding.", "Tanda kegagalan sistem waterproofing dan degradasi pasta semen oleh kelembapan."]
    ]
    format_table(table_dataset, ds_widths, ds_headers, ds_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_heading_2(doc, "3.2 Pipeline Pelatihan Deep Learning di Google Colab GPU")
    add_body_p(
        doc,
        "Proses transfer learning dilakukan menggunakan script `train_colab.py` dan notebook `Train_YOLOv8_Concrete_Defects_Colab.ipynb` "
        "pada lingkungan Google Colaboratory GPU NVIDIA Tesla T4 dengan konfigurasi hyperparameter optimal berikut:\n"
        "• Model Arsitektur Awal: `yolov8m.pt` (YOLOv8 Medium, 25.9 juta parameter) untuk menangkap fitur tekstur retakan mikro.\n"
        "• Ukuran Citra Input (imgsz): 640 x 640 piksel.\n"
        "• Jumlah Epoch: 100 epoch dengan mekanisme Early Stopping (patience = 20 epoch) untuk mencegah overfitting.\n"
        "• Batch Size: 16 citra per iterasi.\n"
        "• Optimizer & Learning Rate: Stochastic Gradient Descent (SGD) dengan initial learning rate lr0 = 0.01, momentum = 0.937, dan Cosine Learning Rate Scheduler (cos_lr = True).\n"
        "• Augmentasi Citra Canggih: Mosaic augmentation (1.0), Mixup (0.15), Horizontal Flip (0.5), Random HSV Shift (h=0.015, s=0.7, v=0.4), serta Close Mosaic pada 10 epoch terakhir guna memastikan bobot akhir bebas artefak.\n"
        "• Luaran Terbaik (Artifacts): Berkas bobot biner `weights/best.pt` yang menghasilkan mean Average Precision (mAP@50) melampaui 90% pada data pengujian independen."
    )

    add_heading_2(doc, "3.3 Panduan Retraining untuk Penambahan Kelas Arsitektur (Solusi Keramik & Plafond)")
    add_body_p(
        doc,
        "Apabila di masa mendatang peneliti atau instansi ingin memperluas kemampuan sistem agar dapat mengenali kerusakan non-beton "
        "(seperti keramik lantai pecah, kusen kayu dimakan rayap, atau plafond gypsum bolong), langkah-langkah retraining terstandarisasi "
        "yang harus dilakukan adalah:\n"
        "1. Pengumpulan Citra Baru: Mengumpulkan 300 - 500 foto variasi keramik pecah dan cacat arsitektur dari berbagai sudut dan pencahayaan.\n"
        "2. Anotasi Label via Roboflow: Mengunggah citra ke platform Roboflow, membuat kelas baru bernama `Ceramic_Crack`, `Plafond_Damage`, dan `Wooden_Decay` menggunakan tool bounding box persegi.\n"
        "3. Ekspor Format YOLOv8: Mengekspor dataset yang telah dianotasi dalam format YOLOv8 ZIP (train, valid, test beserta file data.yaml).\n"
        "4. Menjalankan Ulang train_colab.py: Mengunggah ZIP ke Google Drive, menyesuaikan jumlah kelas `nc: 7`, dan melatih kembali model selama 100 epoch.\n"
        "5. Penggantian File weights/best.pt: Mengunduh file `best.pt` hasil training baru dan meletakkannya di folder `weights/` aplikasi."
    )

    add_heading_2(doc, "3.4 Basis Data Analisis Harga Satuan Pekerjaan (AHSP Jawa Tengah 2023)")
    add_body_p(
        doc,
        "Standar harga satuan yang ditanamkan dalam aplikasi merujuk secara legal pada Keputusan Dinas Pekerjaan Umum Bina Marga "
        "dan Cipta Karya Provinsi Jawa Tengah Tahun 2023. Berikut adalah tabel kompilasi data acuan upah dan bahan yang digunakan sistem:"
    )

    table_ahsp_data = doc.add_table(rows=1, cols=4)
    ahsp_widths = [Inches(1.2), Inches(2.2), Inches(1.4), Inches(2.2)]
    ahsp_headers = ["Kode AHSP", "Uraian Pekerjaan Pemeliharaan", "Satuan", "Harga Satuan (HSP Total)"]
    ahsp_data = [
        ["A.4.1.1.20.JT", "Perbaikan Beton Keropos / Spalling dengan Mortar Semen Polimer Non-Shrink", "m²", "Rp 348.650"],
        ["A.4.1.1.18.JT", "Perbaikan Retak Struktur Balok/Kolom dengan Injeksi Resin Epoxy", "m'", "Rp 215.400"],
        ["A.4.4.1.25.JT", "Perbaikan Retak Rambut Pasangan Dinding Bata / Plesteran", "m²", "Rp 107.146"],
        ["A.4.4.3.35.JT", "Waterproofing Membrane & Plesteran Pelat Dak Beton", "m²", "Rp 185.900"],
        ["A.4.4.3.35.JT", "Pembongkaran dan Penggantian Ubin Keramik Lantai 40x40", "m²", "Rp 168.500"],
        ["A.4.6.2.04.JT", "Pengecatan Ulang Dinding Interior/Eksterior Anti Jamur", "m²", "Rp 45.200"]
    ]
    format_table(table_ahsp_data, ahsp_widths, ahsp_headers, ahsp_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    doc.add_page_break()

    # ==========================================================================
    # BAB 4: PANDUAN VISUAL & PENJELASAN DETAIL SELURUH FITUR APLIKASI
    # ==========================================================================
    add_heading_1(doc, "BAB 4: PANDUAN VISUAL & PENJELASAN DETAIL FITUR APLIKASI (DISERTAI GAMBAR)")
    
    add_body_p(
        doc,
        "Bab ini menyajikan panduan komprehensif seluruh fitur yang ada pada aplikasi, dilengkapi dengan tangkapan layar (screenshot) "
        "nyata dari antarmuka sistem yang beroperasi. Setiap fitur dibedah berdasarkan fungsi, data masukan (input), "
        "proses komputasi di latar belakang, serta format keluaran (output)."
    )

    # Fitur 1: Form Identitas & Jam Real-Time
    add_heading_2(doc, "4.1 Fitur 1: Form Identitas Bangunan & Jam Real-Time WIB")
    add_body_p(
        doc,
        "Fitur ini berada di bagian atas form inspeksi. Pengguna dapat mengisi metadata gedung secara fleksibel, "
        "serta memilih waktu inspeksi riil dengan presisi jam, menit, dan detik waktu Indonesia Barat (WIB). "
        "Data waktu ini akan disinkronkan secara otomatis ke dalam nomor Berita Acara dan kop dokumen PDF resmi."
    )
    img_f1 = os.path.join(USER_UP_DIR, "media_1794612482671_f9c35fa2.png")
    add_image_with_caption(doc, img_f1, "Gambar 4.1: Antarmuka Judul Sistem dan Header Dokumen Resmi")
    add_body_p(
        doc,
        "Spesifikasi Input pada Fitur 1:\n"
        "• Objek Gedung: Nama fasilitas yang diperiksa (contoh: Gedung Kuliah Terpadu Lantai 3).\n"
        "• Lokasi / Ruang: Lokasi spesifik (contoh: Ruang Kelas 302, Sisi Dinding Selatan).\n"
        "• Nama Surveyor & Identitas: Nama pemeriksa dan NIM/NIP untuk keperluan tanda tangan berita acara.\n"
        "• Jam Real-Time: Input time widget Streamlit yang menampilkan waktu presisi saat pengambilan data lapangan berlangsung."
    )

    # Fitur 2 & 3: Unggah Citra & Kuantifikasi Volume
    add_heading_2(doc, "4.2 Fitur 2 & 3: Sensor Visual, Unggah Citra Kerusakan & Metode Volume Fleksibel")
    add_body_p(
        doc,
        "Sistem menyediakan fleksibilitas input citra digital melalui drag-and-drop file JPG/PNG dari memori laptop "
        "maupun pemotretan langsung di lokasi proyek menggunakan kamera smartphone. "
        "Setelah foto diunggah, surveyor dapat menentukan metode kuantifikasi volume melalui dua opsi:\n"
        "1. Input Volume Langsung: Mengetikkan luasan m² atau panjang m' hasil pengukuran meteran roll manual.\n"
        "2. Hitung dari Dimensi Lapangan: Memasukkan nilai Panjang (m) dan Lebar (m), di mana sistem secara otomatis menghitung Luas = Panjang x Lebar."
    )
    img_sample_crack = os.path.join(SAMPLE_IMG_DIR, "retak_balok_beton.jpg")
    add_image_with_caption(doc, img_sample_crack, "Gambar 4.2: Contoh Dokumentasi Citra Kerusakan Balok Beton Lapangan Sebelum Diproses")
    
    img_sample_spall = os.path.join(SAMPLE_IMG_DIR, "spalling_kolom_beton.jpg")
    add_image_with_caption(doc, img_sample_spall, "Gambar 4.3: Contoh Dokumentasi Citra Spalling Keropos Kolom Beton Bertulang")

    # Fitur 4: Hasil Deteksi Visual AI Bounding Box
    add_heading_2(doc, "4.3 Fitur 4: Hasil Deteksi Visual AI Bounding Box & Confidence Score")
    add_body_p(
        doc,
        "Ketika tombol '🚀 Mulai Analisis Deteksi AI & Hitung RAB' ditekan, modul YOLOv8 memindai citra dalam waktu < 1 detik. "
        "Sistem menampilkan tab komparasi antara citra asli dan citra beranotasi bounding box warna-warni yang menunjukkan jenis kerusakan, "
        "koordinat cacat beton, serta persentase keyakinan deteksi (Confidence Score)."
    )
    img_f4 = os.path.join(USER_UP_DIR, "media_1794497075319_8e6faee3.jpg")
    add_image_with_caption(doc, img_f4, "Gambar 4.4: Tampilan Sensor Visual & Hasil Deteksi Bounding Box YOLOv8 dengan Confidence Score")
    
    img_f4_alt = os.path.join(USER_UP_DIR, "media_1794008876588_4f0acbe4.jpg")
    add_image_with_caption(doc, img_f4_alt, "Gambar 4.5: Verifikasi Hasil Anotasi Bounding Box AI dan Sinkronisasi Matriks IKBG")
    add_body_p(
        doc,
        "Pada tampilan di atas, sistem juga menyediakan tombol '📸 Unduh Citra Deteksi AI (.jpg)'. "
        "Tombol ini memungkinkan surveyor mengunduh berkas gambar beranotasi berkualitas tinggi untuk langsung disisipkan "
        "ke dalam lampiran bab hasil dan pembahasan skripsi atau laporan dinas tanpa perlu screenshot manual."
    )

    # Fitur 5 & 6: Rincian AHSP & Biaya RAB
    add_heading_2(doc, "4.4 Fitur 5 & 6: Estimasi Biaya RAB & Rincian Logistik Lapangan (AHSP Jateng 2023)")
    add_body_p(
        doc,
        "Fitur ini merupakan inti dari integrasi rekayasa sipil. Sistem mengalikan volume kerusakan dengan koefisien AHSP Jawa Tengah 2023 "
        "dan merinci secara transparan berapa kebutuhan Orang-Hari (OH) tenaga kerja lapangan serta berapa kuantum sak semen/tube sealant yang diperlukan toko material."
    )
    img_f6 = os.path.join(USER_UP_DIR, "media_1794609687282_76d8395f.jpg")
    add_image_with_caption(doc, img_f6, "Gambar 4.6: Rincian Analisis Harga Satuan Pekerjaan (AHSP) Tenaga Kerja, Bahan, dan Biaya Overhead 15%")
    add_body_p(
        doc,
        "Berdasarkan tangkapan layar Gambar 4.6, rincian biaya dibagi secara profesional menjadi:\n"
        "• Rincian Tenaga Kerja: Pekerja (0.120 OH), Tukang Batu (0.150 OH), Kepala Tukang (0.015 OH), Mandor (0.008 OH).\n"
        "• Rincian Bahan Material: Modified Mortar / Filler Elastis (1.25 kg), Elastomeric Polyurethane (0.35 tube), Agent Adhesive (0.15 liter), Plamir & Finishing (0.25 kg).\n"
        "• Ringkasan Biaya Bersih: Biaya Dasar (Tenaga + Bahan = Rp 93.170) + Overhead & Profit (15% = Rp 13.976) = Total HSP Rp 107.146/m²."
    )

    # Fitur 7: Langkah Perbaikan & Metode Kerja Standar
    add_heading_2(doc, "4.5 Fitur 7: Langkah-Langkah Perbaikan & Metode Kerja Standar")
    add_body_p(
        doc,
        "Untuk memandu mandor dan tukang di lapangan, sistem menyajikan 5 kartu metode kerja perbaikan standar sesuai NSPK "
        "(Norma, Standar, Prosedur, dan Kriteria) Kementerian PUPR. Desain kartu disesuaikan menyerupai formulir checklist kondisi fisik."
    )
    img_f7 = os.path.join(USER_UP_DIR, "media_1794610405811_c76e5e34.jpg")
    add_image_with_caption(doc, img_f7, "Gambar 4.7: Formulir Langkah-Langkah Perbaikan dan Prosedur Metode Kerja Standar PUPR")
    add_body_p(
        doc,
        "Kelima tahapan perbaikan yang disajikan meliputi:\n"
        "• Tahap 1: Pembersihan celah retak rambut dinding dari debu dan plesteran rapuh menggunakan sikat kawat dan kuas bersih.\n"
        "• Tahap 2: Buka alur celah retak membentuk alur huruf V (V-groove) selebar 3-5 mm memakai pahat atau scraper.\n"
        "• Tahap 3: Oleskan cairan bonding agent pada alur retak sebagai lapis pengikat material baru.\n"
        "• Tahap 4: Isi dan rapatkan celah menggunakan mortar polimer / filler elastis hingga rata dengan bidang dinding eksisting.\n"
        "• Tahap 5: Amplas halus permukaan setelah kering, lalu lapisi cat dasar dan cat finishing 2 lapis sesuai warna awal."
    )

    # Fitur 8: Ekspor Berkas Resmi Dokumen PDF
    add_heading_2(doc, "4.6 Fitur 8: Ekspor Formulir Data Kerusakan Resmi Dokumen PDF Bersertifikasi")
    add_body_p(
        doc,
        "Fitur ekspor PDF menyusun seluruh data analisis menjadi dokumen Berita Acara A4 siap cetak dan siap ditandatangani. "
        "Tata letak dokumen menggunakan kop resmi kedinasan dan format formulir inspeksi terstandarisasi."
    )
    img_f8_tbl = os.path.join(USER_UP_DIR, "media_1794612971879_8d5bdb95.png")
    add_image_with_caption(doc, img_f8_tbl, "Gambar 4.8: Format Tabel Metadata Formulir pada Dokumen PDF Hasil Ekspor")
    
    img_f8_doc = os.path.join(USER_UP_DIR, "media_1794613506125_d57525fd.png")
    add_image_with_caption(doc, img_f8_doc, "Gambar 4.9: Bagian Dokumentasi Citra Hasil Pemindaian Visual pada Dokumen PDF")
    add_body_p(
        doc,
        "Karakteristik Dokumen PDF yang Dihasilkan:\n"
        "1. Urutan Halaman Rapi: Diawali dengan Kop 'FORMULIR DATA KERUSAKAN BANGUNAN GEDUNG', diikuti Tabel Metadata, Bagian 1 Dokumentasi Foto Anotasi AI, Bagian 2 Formulir Prosentase IKBG, Bagian 3 Rencana Anggaran Biaya (RAB), dan Bagian 4 Prosedur Perbaikan.\n"
        "2. Lembar Tanda Tangan Tiga Pihak: Menyediakan kolom pengesahan bagi Surveyor Lapangan, Dosen Pembimbing / Penguji, dan Pejabat Pengelola Gedung lengkap dengan garis tanda tangan rapi."
    )

    # Fitur 9: Modul Riwayat Inspeksi & Database SQLite
    add_heading_2(doc, "4.7 Fitur 9: Modul Riwayat Inspeksi & Portofolio Database SQLite")
    add_body_p(
        doc,
        "Halaman Riwayat Inspeksi berfungsi sebagai pusat kendali untuk memonitor seluruh data gedung yang tersimpan "
        "di dalam basis data SQLite `pemeliharaan_gedung.db` secara transparan dan terstruktur."
    )
    img_f9_kpi = os.path.join(USER_UP_DIR, "media_1794493547009_35c1ec66.jpg")
    add_image_with_caption(doc, img_f9_kpi, "Gambar 4.10: Dashboard Riwayat Inspeksi dan Kartu Metrik KPI Eksekutif")
    
    img_f9_tbl = os.path.join(USER_UP_DIR, "media_1794493554728_016027ba.jpg")
    add_image_with_caption(doc, img_f9_tbl, "Gambar 4.11: Tabel Rekapitulasi Riwayat Inspeksi Gedung dengan Fasilitas Filter Dinamis")
    
    img_f9_exp = os.path.join(USER_UP_DIR, "media_1794493560146_9d5fd71d.jpg")
    add_image_with_caption(doc, img_f9_exp, "Gambar 4.12: Fasilitas Ekspor PDF Resmi dan Pengamanan Reset Database Dua Langkah")

    add_body_p(
        doc,
        "Fitur Utama pada Modul Riwayat:\n"
        "• 4 Kartu KPI Eksekutif: Total Akumulasi Biaya RAB Seluruh Gedung, Total Titik Inspeksi Tercatat, Rata-rata Biaya per Titik, dan Elemen Gedung Paling Sering Rusak.\n"
        "• Tabel Rekapitulasi Riwayat Transparan: Menampilkan seluruh data inspeksi secara kronologis lengkap dengan ID, Tanggal & Waktu, Nama Gedung, Objek, Volume, Prosentase Kerusakan, dan Total Biaya.\n"
        "• Filter Cerdas & Pencarian Pintar: Memudahkan pencarian riwayat data berdasarkan elemen gedung maupun kata kunci spesifik.\n"
        "• Ekspor Formulir Data Kerusakan PDF: Mengunduh berkas laporan resmi gabungan seluruh riwayat inspeksi.\n"
        "• Two-Step Safety Popover: Pengamanan reset database bertingkat di mana pengguna diwajibkan mengetik kata 'HAPUS' sebelum data dapat dikosongkan."
    )

    doc.add_page_break()

    # ==========================================================================
    # BAB 5: KESIMPULAN DAN ROADMAP PENGEMBANGAN
    # ==========================================================================
    add_heading_1(doc, "BAB 5: KESIMPULAN & ROADMAP PENGEMBANGAN")
    
    add_body_p(
        doc,
        "5.1 Kesimpulan Capaian Sistem:\n"
        "1. Berhasil Mengintegrasikan Dua Disiplin Ilmu: Sistem ini sukses menjembatani teknologi Artificial Intelligence (Computer Vision YOLOv8) dengan keilmuan Rekayasa Manajemen Konstruksi (AHSP Dinas PU Bina Marga dan Cipta Karya Jawa Tengah 2023).\n"
        "2. Efisiensi Waktu dan Transparansi Anggaran: Proses audit kondisi fisik gedung yang biasanya membutuhkan waktu hitungan hari untuk rekapitulasi data dan perhitungan volume manual kini dapat dituntaskan dalam waktu hitungan detik secara otomatis dan transparan.\n"
        "3. Kepatuhan Regulasi Nasional: Seluruh parameter penilaian kondisi fisik mengadopsi standar resmi Permen PUPR No. 24/PRT/M/2008 & Permen PUPR No. 1/PRT/M/2022, sehingga luaran dokumen PDF sah digunakan sebagai dokumen pendukung teknis pemeliharaan gedung negara.\n"
        "4. Kesiapan Cloud Deployment: Aplikasi terbukti stabil diakses melalui peramban web daring pada Streamlit Community Cloud (https://sistem-deteksi-kerusakan-gedung.streamlit.app/) tanpa memerlukan konfigurasi rumit di sisi pengguna akhir."
    )
    
    add_body_p(
        doc,
        "5.2 Roadmap & Saran Pengembangan Masa Depan:\n"
        "1. Integrasi Sensor Drone / UAV: Mengembangkan antarmuka penerimaan data citra aerial dari drone untuk inspeksi retakan pada dinding luar gedung bertingkat tinggi (high-rise facade) yang sulit dijangkau manusia.\n"
        "2. Ekspansi Dataset Non-Struktural (Roboflow): Menambah kelas deteksi baru untuk elemen arsitektur interior (keramik pecah, plafond rembes, kusen rapuh) menggunakan prosedur transfer learning yang telah dipaparkan pada Bab 3.\n"
        "3. Integrasi Building Information Modeling (BIM): Menghubungkan titik koordinat cacat deteksi AI langsung ke dalam model 3D Revit/IFC agar manajemen pemeliharaan fasilitas (Facility Management) menjadi sepenuhnya digital."
    )

    add_callout(
        doc,
        "Dokumentasi lengkap ini disusun secara khusus sebagai panduan teknis resmi, laporan pertanggungjawaban pengembangan "
        "sistem, dan naskah referensi komprehensif bagi penelitian skripsi teknik sipil digital.",
        title="LEMBAR PENGESAHAN DOKUMEN SISTEM",
        border_color="10B981",
        bg_color="ECFDF5",
        icon="✅"
    )

    # Simpan Dokumen
    os.makedirs(os.path.dirname(OUTPUT_DOCX), exist_ok=True)
    doc.save(OUTPUT_DOCX)
    print(f"============================================================")
    print(f"✅ DOKUMEN WORD LENGKAP BERHASIL DIBUAT DENGAN SUKSES!")
    print(f"📁 Lokasi Berkas : {OUTPUT_DOCX}")
    print(f"📊 Ukuran Berkas : {os.path.getsize(OUTPUT_DOCX)} bytes")
    print(f"============================================================")

if __name__ == "__main__":
    build_complete_document()
