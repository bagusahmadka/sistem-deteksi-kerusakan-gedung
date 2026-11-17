"""
Script Generator Manual Book & Panduan Teknis Sistem Deteksi Kerusakan Gedung (.docx)
Aplikasi Skripsi Teknik Sipil Digital berbasis Deep Learning YOLOv8 & AHSP Jawa Tengah 2023
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    """Mengatur warna background cell tabel."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Mengatur padding dalam cell tabel (dalam dxa, 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout(doc, text, title="PERHATIAN PENTING", border_color="2563EB", bg_color="F0F4FF"):
    """Membuat callout box dengan garis tebal di sebelah kiri."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Border hanya di sebelah kiri
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    
    run_title = p.add_run(f"📌 {title}\n")
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
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_row = tbl.rows[0]
    for i, h_text in enumerate(headers):
        cell = header_row.cells[i]
        cell.width = col_widths[i]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, top=140, bottom=140, left=160, right=160)
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
            elif c_idx == len(row_data) - 1 and any(char.isdigit() for char in str(val)):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def generate_manual_book():
    doc = docx.Document()
    
    # Margin 1 inch (2.54 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    # Gaya Judul Utama (Cover / Heading Utama)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(4)
    run_t1 = p_title.add_run("BUKU PANDUAN PENGGUNAAN & PETUNJUK TEKNIS\n(MANUAL BOOK)")
    run_t1.bold = True
    run_t1.font.name = "Calibri"
    run_t1.font.size = Pt(18)
    run_t1.font.color.rgb = RGBColor(15, 23, 42)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(18)
    run_t2 = p_sub.add_run("SISTEM DETEKSI KERUSAKAN BANGUNAN GEDUNG\nBERBASIS DEEP LEARNING YOLOV8 & ESTIMASI AHSP JAWA TENGAH 2023")
    run_t2.bold = True
    run_t2.font.name = "Calibri"
    run_t2.font.size = Pt(13)
    run_t2.font.color.rgb = RGBColor(37, 99, 235)
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_after = Pt(24)
    run_meta = p_meta.add_run("Aplikasi Skripsi Teknik Sipil Digital | Standar Acuan: Permen PUPR No. 1/PRT/M/2022 & No. 24/PRT/M/2008\nTautan Aplikasi: https://sistem-deteksi-kerusakan-gedung.streamlit.app/")
    run_meta.italic = True
    run_meta.font.name = "Calibri"
    run_meta.font.size = Pt(9.5)
    run_meta.font.color.rgb = RGBColor(71, 85, 105)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # =========================================================================
    # BAB 1: PENDAHULUAN & ARSITEKTUR SISTEM
    # =========================================================================
    h1 = doc.add_heading("BAB 1: PENDAHULUAN & ARSITEKTUR SISTEM", level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(6)
    for run in h1.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(15, 23, 42)
        
    p1 = doc.add_paragraph(
        "Sistem Deteksi Kerusakan Gedung adalah aplikasi website terpadu yang dirancang khusus untuk memodernisasi "
        "tata kelola pemeliharaan dan audit kelayakan fisik bangunan gedung. Aplikasi ini menggabungkan teknologi "
        "Kecerdasan Buatan (Artificial Intelligence / Computer Vision) berbasis Deep Learning YOLOv8 dengan sistem "
        "perhitungan Rencana Anggaran Biaya (RAB) pemeliharaan mengacu pada Analisis Harga Satuan Pekerjaan (AHSP) "
        "Dinas Pekerjaan Umum Bina Marga dan Cipta Karya Provinsi Jawa Tengah Tahun 2023."
    )
    p1.paragraph_format.line_spacing = 1.15
    p1.paragraph_format.space_after = Pt(6)
    
    p2 = doc.add_paragraph(
        "Landasan regulasi dan standardisasi yang diintegrasikan dalam sistem meliputi:\n"
        "1. Peraturan Menteri PUPR No. 1/PRT/M/2022 tentang Pedoman Penyusunan Perkiraan Biaya Pekerjaan Konstruksi Bidang Pekerjaan Umum dan Perumahan Rakyat.\n"
        "2. Peraturan Menteri PUPR No. 24/PRT/M/2008 tentang Pedoman Pemeliharaan dan Perawatan Bangunan Gedung (Standar Indeks Kinerja Bangunan Gedung / IKBG).\n"
        "3. Standar AHSP Dinas PU Bina Marga & Cipta Karya Provinsi Jawa Tengah Tahun 2023."
    )
    p2.paragraph_format.line_spacing = 1.15
    p2.paragraph_format.space_after = Pt(8)

    # Spesifikasi Teknologi
    h2 = doc.add_heading("1.1 Spesifikasi Arsitektur Teknologi", level=2)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    for run in h2.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(30, 58, 138)
        
    table_tech = doc.add_table(rows=1, cols=3)
    tech_widths = [Inches(1.5), Inches(2.2), Inches(2.8)]
    tech_headers = ["Komponen", "Teknologi / Library", "Fungsi dalam Sistem"]
    tech_data = [
        ["Framework Frontend & Backend", "Streamlit (Python)", "Antarmuka responsif (Mobile/PC) & orkestrator data interaktif."],
        ["Model Computer Vision", "YOLOv8 (Ultralytics PyTorch)", "Deteksi objek cacat beton otomatis dari citra visual (best.pt)."],
        ["Database Relasional", "SQLite3", "Penyimpanan persisten seluruh riwayat inspeksi lapangan & log kalkulasi."],
        ["Mesin Ekspor PDF", "ReportLab", "Pencetakan Berita Acara & Laporan Audit resmi A4 lengkap dengan kop PU & lembar tanda tangan."],
        ["Mesin Ekspor Excel", "OpenPyXL", "Pembuatan formulir evaluasi kondisi fisik & fungsi (IKBG) beserta rincian lembar RAB."]
    ]
    format_table(table_tech, tech_widths, tech_headers, tech_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # =========================================================================
    # BAB 2: BATASAN & KARAKTERISTIK MODEL AI (SANGAT KRUSIAL)
    # =========================================================================
    h1 = doc.add_heading("BAB 2: BATASAN & KARAKTERISTIK MODEL AI YOLOV8 (PENTING)", level=1)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    for run in h1.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(15, 23, 42)

    p_limit_intro = doc.add_paragraph(
        "Model Deep Learning YOLOv8 dilatih secara spesifik menggunakan dataset cacat permukaan beton dan elemen struktural. "
        "Agar model memberikan hasil deteksi yang akurat, berdaya guna tinggi, dan tidak menghasilkan prediksi asal-asalan (halusinasi/false positive), "
        "surveyor wajib memahami batasan kelas terlatih dan objek apa saja yang TIDAK BISA dideteksi oleh sistem."
    )
    p_limit_intro.paragraph_format.line_spacing = 1.15
    p_limit_intro.paragraph_format.space_after = Pt(6)

    h2 = doc.add_heading("2.1 Empat (4) Kelas Kerusakan yang Dikenali Model YOLOv8", level=2)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    for run in h2.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(30, 58, 138)

    table_classes = doc.add_table(rows=1, cols=4)
    cls_widths = [Inches(0.6), Inches(1.3), Inches(2.2), Inches(2.4)]
    cls_headers = ["No", "Nama Kelas", "Karakteristik Visual", "Kode AHSP & Rekomendasi Terkait"]
    cls_data = [
        ["1", "Crack\n(Retakan)", "Garis retak memanjang pada permukaan plesteran dinding, balok, atau pelat beton.", "A.4.4.1.25.JT (Retak Dinding) / A.4.1.1.18.JT (Retak Struktur Balok dengan Injeksi Epoxy)."],
        ["2", "Spalling\n(Beton Keropos)", "Permukaan selimut beton terkelupas, rontok, pecah, atau berlubang meninggalkan cekungan kasar.", "A.4.1.1.20.JT (Patching Mortar Semen Polimer Struktural Non-Shrink)."],
        ["3", "Exposed Rebar\n(Besi Tulangan Terbuka)", "Besi tulangan baja di dalam beton terlihat keluar akibat selimut beton terlepas, seringkali disertai karat.", "A.4.1.1.20.JT (Pembersihan karat, aplikasi zinc-rich primer, dan pelapisan mortar polimer)."],
        ["4", "Efflorescence\n(Pengapuran / Garam)", "Bercak putih garam kapur atau noda rembesan air pada dak beton, dinding bawah, atau sambungan pelat.", "A.4.4.3.35.JT (Waterproofing Membrane Pelat Dak) / A.4.6.2.04.JT (Cat Anti Alkali & Jamur)."]
    ]
    format_table(table_classes, cls_widths, cls_headers, cls_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Objek yang TIDAK BISA dideteksi
    h2 = doc.add_heading("2.2 Bahan & Kondisi yang TIDAK BISA Dideteksi Model (Hindari Pemotretan)", level=2)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    for run in h2.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(30, 58, 138)

    p_avoid = doc.add_paragraph(
        "Model ini HANYA dilatih untuk mengenali cacat material berbasis BETON dan PASANGAN DINDING. "
        "Berikut adalah daftar material, elemen, dan anomali visual yang TIDAK BOLEH difoto karena tidak masuk ke dalam kelas pelatihan dan dapat menyebabkan deteksi keliru (False Positive / Misclassification):"
    )
    p_avoid.paragraph_format.line_spacing = 1.15
    p_avoid.paragraph_format.space_after = Pt(6)

    table_avoid = doc.add_table(rows=1, cols=3)
    av_widths = [Inches(1.8), Inches(2.3), Inches(2.4)]
    av_headers = ["Kategori Objek / Bahan", "Contoh Kondisi Lapangan", "Dampak Jika Tetap Difoto"]
    av_data = [
        ["Elemen Non-Beton (Kayu, Kaca, Baja)", "Kusen kayu lapuk/dimakan rayap, kaca jendela retak/pecah, atap seng berkarat, rangka plafon baja ringan bengkok.", "Model tidak mengenali material kayu/kaca; model mungkin keliru menebak retak kaca sebagai 'Crack' beton atau tidak mendeteksi sama sekali."],
        ["Anomali Visual Mirip Garis Retak", "Kabel listrik tipis warna hitam/abu menempel di dinding, tali kawat, sambungan nat keramik, alur tali air arsitektur.", "Model rawan mengira kabel panjang atau nat keramik sebagai 'Crack' (False Positive)."],
        ["Bayangan Tajam & Pencahayaan Ekstrem", "Bayangan ranting pohon atau tiang yang tajam melintang di dinding akibat terik matahari.", "Kontras bayangan hitam tipis sering disalahartikan sebagai garis retakan."],
        ["Noda Cairan & Kotoran Permukaan", "Noda tumpahan kopi/oli, coretan spidol/grafiti, kotoran lumpur kering, atau sarang laba-laba tebal.", "Model bisa salah mengklasifikasikan bercak cat/kotoran sebagai 'Efflorescence' atau 'Spalling'."],
        ["Lumut Tebal & Jamur Hitam Basah", "Lapisan lumut hijau basah tebal yang menutup total tekstur beton di saluran air luar gedung.", "Menghalangi tekstur retakan asli beton sehingga bounding box tidak dapat mengunci area kerusakan."]
    ]
    format_table(table_avoid, av_widths, av_headers, av_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_callout(
        doc,
        "1. Pastikan objek yang difoto adalah elemen struktural atau arsitektural berbasis beton/dinding (balok, kolom, pelat lantai, dinding bata/plesteran).\n"
        "2. Singkirkan kabel yang menggelantung di depan retakan sebelum memotret.\n"
        "3. Jangan memotret kerusakan kayu, kusen aluminium penyok, kaca jendela pecah, atau plafon gypsum bolong karena bukan target latih model YOLOv8 ini.",
        title="ATURAN EMAS PEMOTRETAN LAPANGAN (GOLDEN RULES)",
        border_color="DC2626",
        bg_color="FEF2F2"
    )

    # SOP Pemotretan
    h2 = doc.add_heading("2.3 Standar Operasional Pemotretan Lapangan (SOP Photography)", level=2)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    for run in h2.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(30, 58, 138)

    p_sop = doc.add_paragraph(
        "Untuk menjamin nilai keyakinan (confidence score) model di atas 80% dan hasil estimasi RAB akurat:\n"
        "1. Jarak Pemotretan: Ambil foto pada jarak 0.5 meter hingga 1.5 meter tegak lurus (90 derajat) terhadap bidang kerusakan.\n"
        "2. Pencahayaan Cukup: Hindari kondisi remang-remang. Pada area basement atau kolong tangga, gunakan bantuan senter atau flash ponsel.\n"
        "3. Ketajaman & Fokus: Pastikan kamera terkunci pada tekstur retakan. Jangan bergerak saat menekan tombol rana untuk menghindari motion blur.\n"
        "4. Komposisi Citra: Posisikan area kerusakan berada di tengah bidang foto (centered) dan mencakup sebagian bidang beton sehat sebagai pembanding latar."
    )
    p_sop.paragraph_format.line_spacing = 1.15
    p_sop.paragraph_format.space_after = Pt(12)

    # =========================================================================
    # BAB 3: PENJELASAN LENGKAP FITUR-FITUR APLIKASI
    # =========================================================================
    h1 = doc.add_heading("BAB 3: PENJELASAN SELURUH FITUR APLIKASI", level=1)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    for run in h1.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(15, 23, 42)

    p_f_intro = doc.add_paragraph(
        "Aplikasi terbagi menjadi tiga halaman navigasi utama di bilah samping (sidebar) dengan fungsi yang terintegrasi secara modular:"
    )
    p_f_intro.paragraph_format.space_after = Pt(6)

    # Halaman 1
    h2 = doc.add_heading("3.1 Modul Halaman 1: Dashboard Utama & Input Data", level=2)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    for run in h2.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(30, 58, 138)

    p_h1 = doc.add_paragraph(
        "Halaman utama berfungsi sebagai pusat operasional inspeksi kerusakan gedung. Fitur-fitur yang terdapat di halaman ini antara lain:\n"
        "• Accordion Panduan Alur 3 Langkah: Ringkasan alur kerja cepat bagi penguji/pengguna baru.\n"
        "• Identitas Bangunan & Lembar Pengesahan Laporan: Input nama gedung, lokasi/lantai, nama surveyor & NIM/NIP, dosen pembimbing & NIP, serta pengelola gedung & NIP untuk dicetak otomatis pada berkas PDF resmi.\n"
        "• Modul Unggah Foto Kerusakan: Menerima format JPG, JPEG, dan PNG dari penyimpanan laptop maupun pemotretan langsung kamera HP.\n"
        "• Pemilih Elemen Gedung: Pilihan 8 komponen standar (Balok, Kolom, Dinding Bata, Pelat Lantai, Plafon, Keramik, Kusen, Dinding Finishing).\n"
        "• Metode Penentuan Volume Fleksibel: Pilihan antara 'Input Volume Langsung' atau 'Hitung dari Dimensi Lapangan (Panjang x Lebar)'.\n"
        "• Sensor Visual & Tab Anotasi AI: Menampilkan foto lapangan asli dan foto hasil pemindaian bounding box YOLOv8 lengkap dengan label kerusakan dan confidence score.\n"
        "• Tombol Unduh Citra Deteksi AI (.jpg): Tombol satu klik untuk mengunduh gambar berkotak AI resolusi asli guna disisipkan pada Bab 4 Skripsi.\n"
        "• Formulir Penilaian Kondisi Fisik & Fungsi (IKBG PUPR): Perhitungan skor fisik (50%) dan fungsi (50%) serta penentuan klasifikasi tindakan (RUTIN, PREVENTIF, KOREKTIF, REHABILITATIF) sesuai rumus kinerja µku.\n"
        "• Rincian RAB Standar AHSP Jateng 2023: Tabel perkalian Volume x Harga Satuan Pekerjaan (HSP).\n"
        "• Breakdown Kebutuhan Riil Logistik Lapangan: Menghitung kebutuhan Orang-Hari (OH) Tenaga Kerja (Pekerja, Tukang, Kepala Tukang, Mandor) dan kuantum Bahan Material (kg semen polimer, tube sealant, liter bonding agent).\n"
        "• Slider Simulasi Cepat Penyesuaian Volume (What-If Scenario): Slider interaktif untuk menguji perubahan volume kerusakan secara langsung tanpa perlu upload ulang.\n"
        "• Ekspor Berita Acara Tunggal (PDF ReportLab & Excel OpenPyXL): Menghasilkan berkas cetak siap tanda tangan dan dokumen spreadsheet lengkap."
    )
    p_h1.paragraph_format.line_spacing = 1.15
    p_h1.paragraph_format.space_after = Pt(8)

    # Halaman 2
    h2 = doc.add_heading("3.2 Modul Halaman 2: Riwayat Inspeksi Gedung & Portofolio", level=2)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    for run in h2.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(30, 58, 138)

    p_h2 = doc.add_paragraph(
        "Halaman riwayat berfungsi untuk memonitor seluruh data hasil inspeksi yang tersimpan di database SQLite:\n"
        "• Kartu Eksekutif IKBG Global: Menganalisis rata-rata skor kesehatan seluruh gedung (µ̄ku) dan memberikan status keandalan serta rekomendasi manajerial institusional.\n"
        "• 4 Kartu Metrik KPI: Total RAB Akumulatif, Total Titik Inspeksi, Rata-rata Biaya per Kerusakan, dan Elemen Gedung Paling Rentan.\n"
        "• Grafik Visual Interaktif (Altair Chart): Diagram batang alokasi anggaran per elemen dan diagram donat proporsi tingkat keparahan (Ringan, Sedang, Berat).\n"
        "• Filter Dinamis & Kotak Pencarian Pintar: Memfilter data berdasarkan elemen bangunan atau kata kunci tertentu.\n"
        "• Tabel Rekapitulasi Terstruktur: Menampilkan seluruh data inspeksi dengan scroll horizontal yang nyaman diakses via HP maupun laptop.\n"
        "• Baris Tombol Ekspor Terpadu: Ekspor Excel (.xlsx), Ekspor PDF Resmi Gabungan (ReportLab), dan Ekspor Database CSV.\n"
        "• Pengamanan Reset Database Dua Langkah (Two-Step Safety Popover): Mencegah penghapusan tak sengaja dengan mewajibkan pengetikan kata kunci 'HAPUS'.\n"
        "• Fitur Hapus Catatan per ID: Memungkinkan penghapusan data spesifik yang keliru tanpa memengaruhi catatan lainnya."
    )
    p_h2.paragraph_format.line_spacing = 1.15
    p_h2.paragraph_format.space_after = Pt(8)

    # Halaman 3
    h2 = doc.add_heading("3.3 Modul Halaman 3: Panduan & Katalog AHSP Jawa Tengah 2023", level=2)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    for run in h2.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(30, 58, 138)

    p_h3 = doc.add_paragraph(
        "Halaman katalog menyediakan transparansi basis data acuan hukum teknis:\n"
        "• Penelusuran Katalog AHSP: Daftar lengkap analisa harga satuan pekerjaan bidang ke-PU-an Cipta Karya Jawa Tengah.\n"
        "• Filter Kategori Pekerjaan: Pekerjaan Struktur Beton, Dinding & Finishing, Waterproofing, Pengecatan, dll.\n"
        "• Rincian Koefisien & Prosedur SOP: Membuka transparansi koefisien upah, koefisien bahan, dan 5 tahapan prosedur perbaikan standar PUPR.\n"
        "• Kalkulator Estimasi Cepat AHSP: Fasilitas bagi surveyor untuk menghitung taksiran biaya pekerjaan tertentu secara mandiri."
    )
    p_h3.paragraph_format.line_spacing = 1.15
    p_h3.paragraph_format.space_after = Pt(12)

    # =========================================================================
    # BAB 4: PANDUAN LANGKAH DEMI LANGKAH PENGGUNAAN APLIKASI
    # =========================================================================
    h1 = doc.add_heading("BAB 4: PANDUAN PENGGUNAAN LANGKAH DEMI LANGKAH (TUTORIAL)", level=1)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    for run in h1.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(15, 23, 42)

    p_tut_intro = doc.add_paragraph(
        "Berikut adalah panduan operasional standar (SOP) pelaksanaan inspeksi menggunakan aplikasi:"
    )
    p_tut_intro.paragraph_format.space_after = Pt(4)

    steps = [
        ("Langkah 1: Membuka Aplikasi & Menyiapkan Identitas", 
         "Akses tautan https://sistem-deteksi-kerusakan-gedung.streamlit.app/ melalui peramban Google Chrome di laptop atau smartphone.\n"
         "Buka expander '🏢 Identitas Bangunan & Lembar Pengesahan Laporan'. Masukkan nama gedung (misal: Gedung Kuliah Terpadu), lokasi spesifik (misal: Lantai 2 - Ruang 204), nama surveyor beserta NIM/NIP, dosen pembimbing, dan nama pejabat pengelola gedung."),
        ("Langkah 2: Mengunggah Dokumentasi Citra Kerusakan", 
         "Klik tombol 'Browse files' pada modul '1. Unggah Foto Kerusakan'.\n"
         "• Jika menggunakan laptop: Pilih file foto JPEG/PNG yang sudah dipersiapkan.\n"
         "• Jika menggunakan smartphone: Anda dapat memilih galeri atau langsung memilih opsi 'Kamera' untuk memotret dinding/balok retak di tempat."),
        ("Langkah 3: Memilih Elemen & Menentukan Besaran Volume", 
         "Pilih elemen gedung yang sedang diperiksa (misal: 'Dinding Pasangan Bata / Hebel' atau 'Balok Beton Bertulang').\n"
         "Pilih metode penentuan volume:\n"
         "• Opsi A (Input Langsung): Ketikkan volume kerusakan (misal: 14.00 m²).\n"
         "• Opsi B (Hitung Dimensi Lapangan): Masukkan Panjang (misal: 7.0 meter) dan Lebar (misal: 2.0 meter), sistem akan otomatis menghitung Luas = 14.00 m²."),
        ("Langkah 4: Menjalankan Analisis Deteksi AI", 
         "Klik tombol besar berwarna biru: '🚀 Mulai Analisis Deteksi AI & Hitung RAB'.\n"
         "Sistem akan memproses citra melalui model YOLOv8 dalam waktu 1-3 detik. Muncul kotak status hijau resmi: '🟢 Status Engine: YOLOv8 Computer Vision Engine Aktif & Optimal'."),
        ("Langkah 5: Memeriksa Hasil Deteksi & Estimasi Biaya", 
         "Periksa hasil analisis pada layar:\n"
         "1. Bounding Box AI: Lihat kotak deteksi pada tab sensor visual, lalu klik tombol '📸 Unduh Citra Deteksi AI (.jpg)' jika ingin menyimpan foto beranotasi.\n"
         "2. Matriks IKBG: Tinjau skor fisik & fungsi serta status klasifikasi (Rutin/Preventif/Korektif/Rehabilitatif).\n"
         "3. Breakdown Kebutuhan Lapangan: Lihat rincian jumlah Orang-Hari (OH) tukang yang dibutuhkan serta volume sak semen/tube sealant yang harus disiapkan toko material.\n"
         "4. Prosedur Kerja: Baca rekomendasi teknis perbaikan sesuai Permen PUPR No. 24/2008."),
        ("Langkah 6: Mengunduh Dokumen Laporan Resmi (PDF & Excel)", 
         "Gulir ke bawah pada bagian tombol ekspor:\n"
         "• Klik tombol hijau '📊 Unduh Formulir Penilaian & RAB (Excel .xlsx)' untuk mengunduh lembar kerja spreadsheet.\n"
         "• Klik tombol biru '📄 Unduh Laporan Resmi & Penilaian (PDF)' untuk mengunduh berkas Berita Acara resmi ReportLab yang memuat Kop PU, hasil foto anotasi, dan tanda tangan lengkap yang telah Anda input.")
    ]

    for title_s, desc_s in steps:
        h2 = doc.add_heading(title_s, level=2)
        h2.paragraph_format.space_before = Pt(8)
        h2.paragraph_format.space_after = Pt(2)
        for run in h2.runs:
            run.font.name = "Calibri"
            run.font.color.rgb = RGBColor(30, 58, 138)
        p_desc = doc.add_paragraph(desc_s)
        p_desc.paragraph_format.line_spacing = 1.15
        p_desc.paragraph_format.space_after = Pt(6)

    # =========================================================================
    # BAB 5: TANYA JAWAB TEKNIS & TROUBLESHOOTING (FAQ)
    # =========================================================================
    h1 = doc.add_heading("BAB 5: TANYA JAWAB TEKNIS & TROUBLESHOOTING (FAQ)", level=1)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    for run in h1.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(15, 23, 42)

    faqs = [
        ("Q1: Mengapa model tidak mendeteksi kotak pada foto yang saya unggah?",
         "Penyebab utama: (a) Objek yang difoto bukan termasuk dalam 4 kelas latih (Crack, Spalling, Exposed Rebar, Efflorescence), (b) Foto terlalu buram/blur, atau (c) Tingkat keyakinan berada di bawah ambang batas deteksi. Sistem secara cerdas akan mengalihkan ke mode acuan elemen terstandarisasi sehingga estimasi RAB tetap dapat dihitung secara valid."),
        ("Q2: Apakah saya harus selalu mengisi NIP Dosen Pembimbing dan Pengelola Gedung?",
         "Tidak wajib (bersifat opsional). Jika Anda mengisinya, nama dan NIP akan langsung tercetak rapi di lembar PDF resmi. Jika dikosongkan, dokumen PDF tetap mencetak format garis titik-titik rapi untuk keperluan tanda tangan basah manual menggunakan pulpen."),
        ("Q3: Mengapa tombol Reset Database tidak langsung menghapus data saat diklik?",
         "Ini adalah fitur keamanan (Two-Step Safety Popover) untuk melindungi data hasil inspeksi Anda agar tidak terhapus secara tidak sengaja saat presentasi di depan dosen penguji. Anda diwajibkan mengetik kata 'HAPUS' dengan benar sebelum database dapat dikosongkan."),
        ("Q4: Bagaimana cara menjalankan aplikasi secara 100% offline saat sidang tanpa koneksi internet?",
         "Jalankan file 'run_app.bat' di laptop Anda. Nyalakan Hotspot portabel dari smartphone Anda (tanpa kuota pun bisa) dan hubungkan laptop ke hotspot tersebut. Buka browser HP dan akses alamat IP lokal laptop (http://IP-LAPTOP:8501). Sistem akan berjalan 100% offline mulus di HP Anda.")
    ]

    for q, a in faqs:
        p_faq = doc.add_paragraph()
        p_faq.paragraph_format.space_before = Pt(6)
        p_faq.paragraph_format.space_after = Pt(2)
        p_faq.paragraph_format.line_spacing = 1.15
        run_q = p_faq.add_run(q + "\n")
        run_q.bold = True
        run_q.font.color.rgb = RGBColor(30, 58, 138)
        run_a = p_faq.add_run(a)
        run_a.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(14)
    p_closing = doc.add_paragraph()
    p_closing.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_close = p_closing.add_run("--- DOKUMEN MANUAL BOOK RESMI SISTEM DETEKSI KERUSAKAN GEDUNG ---")
    run_close.bold = True
    run_close.font.size = Pt(9)
    run_close.font.color.rgb = RGBColor(148, 163, 184)

    output_path = os.path.join(r"d:\gedung", "MANUAL_BOOK_SISTEM_DETEKSI_KERUSAKAN_GEDUNG.docx")
    doc.save(output_path)
    print(f"Manual Book successfully generated at: {output_path}")

if __name__ == "__main__":
    generate_manual_book()
