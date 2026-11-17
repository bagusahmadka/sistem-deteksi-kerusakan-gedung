"""
Mock AI Agent Antigravity Server
Server simulasi independen untuk menerima HTTP POST call dari aplikasi Streamlit
Menjalankan endpoint: POST http://localhost:8000/api/analyze-damage
"""

import http.server
import socketserver
import json
import base64
import random
from datetime import datetime

PORT = 8000

# Database referensi pengetahuan teknik sipil AHSP Jawa Tengah 2023
AHSP_KNOWLEDGE_BASE = {
    "Dinding Pasangan Bata / Hebel": {
        "jenis_kerusakan": "Retak Rambut Plesteran (> 0.20 mm)",
        "tingkat_keparahan": "Sedang",
        "kode_ahsp": "A.4.4.1.25.JT",
        "deskripsi_ahsp": "Perbaikan Retak Plesteran Mortar Polimer & Sealant Elastomeric",
        "satuan": "m2",
        "harga_satuan": 88500,
        "langkah_kerja": [
            "Pembersihan celah retakan dari debu dan kotoran dengan sikat kawat dan kompresor udara.",
            "Pelebaran alur retakan membentuk pola V-groove sedalam 5-10 mm sepanjang retakan.",
            "Aplikasi bonding agent / cairan primer adhesi pada celah retak untuk merekatkan spesi baru.",
            "Pengisian pasta polymer modified mortar / elastomeric sealant secara merata dan dipadatkan.",
            "Finishing plamir khusus retak dan pengecatan ulang 2 lapis sesuai warna eksisting."
        ],
        "rekomendasi": "Pasang crack monitoring gauge (crack meter) selama 30 hari untuk memverifikasi kestabilan struktur. Cegah rembesan air dari talang eksterior."
    },
    "Balok Beton Bertulang": {
        "jenis_kerusakan": "Retak Diagonal Lentur/Geser Balok Struktur",
        "tingkat_keparahan": "Berat",
        "kode_ahsp": "A.4.1.1.18.JT",
        "deskripsi_ahsp": "Perbaikan Beton Retak Struktural dengan Injeksi Epoxy Resin Bertekanan",
        "satuan": "m'",
        "harga_satuan": 315000,
        "langkah_kerja": [
            "Pemetaan dan pengukuran lebar celah retak menggunakan crack detection microscope.",
            "Pembersihan permukaan beton dan penutupan celah permukaan luar menggunakan surface sealer epoxy.",
            "Pemasangan nipple / injection port pada interval 15 - 20 cm sepanjang garis retak.",
            "Injeksi resin epoxy viskositas rendah menggunakan pompa tekanan rendah berkelanjutan hingga keluar dari port berikutnya.",
            "Pemotongan nipple port dan perataan permukaan beton setelah curing 24 jam."
        ],
        "rekomendasi": "Lakukan evaluasi kapasitas dukung balok terhadap beban layan eksisting. Batasi beban hidup lantai di atas balok sementara waktu."
    },
    "Kolom Struktur Beton": {
        "jenis_kerusakan": "Spalling / Keropos Beton & Selimut Terkelupas",
        "tingkat_keparahan": "Berat",
        "kode_ahsp": "A.4.1.1.20.JT",
        "deskripsi_ahsp": "Patching Mortar Semen Polimer Struktural + Anti Karat Tulangan",
        "satuan": "m2",
        "harga_satuan": 245000,
        "langkah_kerja": [
            "Chipping / pembongkaran beton keropos hingga batas beton padat (sound concrete).",
            "Pembersihan karat pada baja tulangan menggunakan sikat baja kawat (wire brush) s.d. standar Sa 2.0.",
            "Aplikasi lapisan pelindung anti korosi (zinc rich primer / rust inhibitor) pada baja tulangan.",
            "Penyapuan bonding agent sintetis pada bidang kontak beton lama.",
            "Aplikasi mortar polimer non-shrink (patching mortar) lapis demi lapis dan pemadatan sempurna."
        ],
        "rekomendasi": "Pastikan tebal selimut beton minimal 30 mm terpenuhi kembali. Lakukan uji Non-Destructive Test (UPV Test) untuk memastikan tidak ada delaminasi internal."
    },
    "Pelat Lantai / Dak Beton": {
        "jenis_kerusakan": "Rembesan Air & Retak Rambut Plat Atap Dak",
        "tingkat_keparahan": "Sedang",
        "kode_ahsp": "A.4.4.3.35.JT",
        "deskripsi_ahsp": "Aplikasi Waterproofing Polyurethane / Membran Bakar Pelat Atap",
        "satuan": "m2",
        "harga_satuan": 175000,
        "langkah_kerja": [
            "Pembersihan lumut, debu, dan genangan air pada permukaan plat dak beton.",
            "Perbaikan retak susut menggunakan kristalisasi semen / mastic sealant elastis.",
            "Aplikasi primer penetrasi bitumen / epoxy primer pada seluruh area plat dak dan fillet sudut chamfer.",
            "Penggelaran lapisan waterproofing membran tahan sinar UV (2 lapis bersilang).",
            "Uji rendam air (ponding test) selama minimal 24 jam untuk verifikasi kekedapan air."
        ],
        "rekomendasi": "Perbaiki kemiringan (slope) screed atap minimal 1-2% menuju lubang drainase (roof drain) agar tidak terjadi genangan air (ponding)."
    },
    "Plafon Gypsum / Akustik": {
        "jenis_kerusakan": "Plafon Lendut, Lapuk & Noda Coklat Akibat Bocor",
        "tingkat_keparahan": "Ringan",
        "kode_ahsp": "A.4.5.1.07.JT",
        "deskripsi_ahsp": "Penggantian Penutup Plafon Gypsum Board 9 mm Rangka Galvanis",
        "satuan": "m2",
        "harga_satuan": 135000,
        "langkah_kerja": [
            "Pelepasan papan gypsum yang telah lapuk dan rapuh dengan hati-hati.",
            "Inspeksi dan perbaikan sumber kebocoran instalasi pipa plumbing atau atap di atas plafon.",
            "Penguatan dan perataan level gantungan kawat hollow galvanis.",
            "Pemasangan lembaran gypsum board baru dengan sekrup drywall berjarak maksimal 20 cm.",
            "Perekatan paper tape pada sambungan, pendempulan compound, amplas halus, dan pengecatan."
        ],
        "rekomendasi": "Pastikan kebocoran pada plat dak atap atau instalasi sanitair di atas plafon telah dihentikan total sebelum memasang lembaran gypsum baru."
    },
    "Lantai Keramik / Granit": {
        "jenis_kerusakan": "Keramik Menggelembung & Terangkat (Popping)",
        "tingkat_keparahan": "Sedang",
        "kode_ahsp": "A.4.4.3.02.JT",
        "deskripsi_ahsp": "Pembongkaran & Pasang Ulang Keramik Lantai Perekat Semen Instan",
        "satuan": "m2",
        "harga_satuan": 220000,
        "langkah_kerja": [
            "Pembongkaran keramik yang berbunyi kopong dan terangkat dengan pahat pipih.",
            "Pembersihan spesi adukan lama hingga ke dasar plat lantai beton.",
            "Aplikasi adukan semen instan mortar perekat ubin (tile adhesive) dengan roskam bergerigi.",
            "Pemasangan ubin keramik baru dengan nat ekspansi selebar 2-3 mm.",
            "Pengisian nat menggunakan tile grout polimer kedap air setelah 24 jam pemasangan."
        ],
        "rekomendasi": "Sediakan expansion joint pada perimeter ruangan luas untuk mengantisipasi muai-susut termal ubin keramik."
    }
}

DEFAULT_FALLBACK = {
    "jenis_kerusakan": "Pelapukan Dinding & Pengelupasan Cat Bangunan",
    "tingkat_keparahan": "Ringan",
    "kode_ahsp": "A.4.7.1.10.JT",
    "deskripsi_ahsp": "Pengerokan Cat Lama & Pengecatan Ulang Dinding Gedung 2 Lapis",
    "satuan": "m2",
    "harga_satuan": 46500,
    "langkah_kerja": [
        "Pengerokan cat lama yang mengelupas dan pengapuran dengan kape besi.",
        "Pencucian dinding dan sterilisasi anti jamur / lumut.",
        "Aplikasi alkali resisting primer sealer 1 lapis untuk netralisasi pH dinding.",
        "Pendempulan tipis pada pori-pori permukaan menggunakan plamir acrylic.",
        "Pengecatan cat dinding akrilik bermutu tinggi minimal 2 lapis hingga menutup sempurna."
    ],
    "rekomendasi": "Perbaiki kelembapan dinding dan ventilasi ruangan untuk menjaga keawetan lapisan cat baru."
}


class MockAIAgentHandler(http.server.BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        self.end_headers()

    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        info = {
            "agent_name": "Antigravity Civil Vision Agent",
            "version": "2.4.0",
            "standard": "AHSP Jawa Tengah 2023",
            "status": "online",
            "endpoint": "/api/analyze-damage"
        }
        self.wfile.write(json.dumps(info, indent=2).encode('utf-8'))

    def do_POST(self):
        if self.path == "/api/analyze-damage" or self.path.startswith("/api/"):
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)

            try:
                request_json = json.loads(post_data.decode('utf-8'))
            except Exception:
                request_json = {}

            elemen_input = request_json.get("elemen_gedung", "Dinding Pasangan Bata / Hebel")
            volume_input = float(request_json.get("volume", 1.0))
            filename = request_json.get("filename", "damage_photo.jpg")

            # Pilih data teknik sipil yang sesuai
            selected = None
            for key in AHSP_KNOWLEDGE_BASE:
                if key.lower() in elemen_input.lower() or elemen_input.lower() in key.lower():
                    selected = AHSP_KNOWLEDGE_BASE[key]
                    elemen_matched = key
                    break

            if not selected:
                elemen_matched = elemen_input if elemen_input else "Dinding Pasangan Bata / Hebel"
                selected = DEFAULT_FALLBACK

            harga_satuan = selected["harga_satuan"]
            total_biaya = int(round(volume_input * harga_satuan))
            confidence = round(random.uniform(94.5, 98.8), 1)

            response_data = {
                "status": "success",
                "agent_id": "antigravity-civil-agent-jateng-v2",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "filename": filename,
                "elemen_gedung": elemen_matched,
                "jenis_kerusakan": selected["jenis_kerusakan"],
                "tingkat_keparahan": selected["tingkat_keparahan"],
                "confidence_score": confidence,
                "kode_ahsp": selected["kode_ahsp"],
                "deskripsi_ahsp": selected["deskripsi_ahsp"],
                "volume": volume_input,
                "satuan": selected["satuan"],
                "harga_satuan": harga_satuan,
                "total_biaya": total_biaya,
                "langkah_kerja": selected["langkah_kerja"],
                "rekomendasi": selected["rekomendasi"],
                "acuan_regulasi": "AHSP Dinas PU Bina Marga & Cipta Karya Jawa Tengah 2023"
            }

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response_data, indent=2).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'{"error": "Endpoint not found"}')

    def log_message(self, format, *args):
        # Override untuk logging bersih
        print(f"[MOCK AI AGENT] {args[0]} - {args[1]}")


def run_server():
    server_address = ('', PORT)
    with socketserver.TCPServer(server_address, MockAIAgentHandler) as httpd:
        print(f"================================================================")
        print(f" Antigravity Civil Vision AI Agent Mock Server Berjalan...")
        print(f" Port: {PORT}")
        print(f" URL Endpoint: http://localhost:{PORT}/api/analyze-damage")
        print(f" Siap menerima HTTP POST request dari Streamlit!")
        print(f"================================================================")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer dihentikan.")


if __name__ == "__main__":
    run_server()
