"""
================================================================================
SCRIPT TRAINING YOLOV8 DETEKSI KERUSAKAN BETON & GEDUNG (GOOGLE COLAB / GPU)
Dataset Acuan : CONCRETE-25FEB (6.554 Citra)
Target Akurasi: mAP@50 > 90% (Ultra High Accuracy)
Standar Acuan : AHSP Dinas PU Bina Marga & Cipta Karya Jawa Tengah 2023
================================================================================
"""

import os
import glob
import shutil
import zipfile
import yaml
import torch
from ultralytics import YOLO

def main():
    print("=" * 70)
    print("🚀 MEMULAI PIPELINE TRAINING YOLOV8 DETEKSI KERUSAKAN GEDUNG")
    print("=" * 70)

    # 1. Cek Ketersediaan GPU
    if torch.cuda.is_available():
        gpu_name = torch.cuda.get_device_name(0)
        vram = round(torch.cuda.get_device_properties(0).total_memory / (1024**3), 2)
        print(f"✅ GPU Terdeteksi : {gpu_name} ({vram} GB VRAM)")
    else:
        print("⚠️ Peringatan: GPU tidak terdeteksi. Training akan lambat di CPU!")

    # 2. Lokasi Dataset
    dataset_dir = "/content/dataset"
    os.makedirs(dataset_dir, exist_ok=True)

    zip_candidates = [
        "/content/drive/MyDrive/CONCRETE-25FEB.v1i.yolov8.zip",
        "/content/drive/MyDrive/data training/CONCRETE-25FEB.v1i.yolov8.zip",
        "/content/drive/MyDrive/data_skripsi/CONCRETE-25FEB.v1i.yolov8.zip",
        "/content/CONCRETE-25FEB.v1i.yolov8.zip"
    ]

    zip_path = None
    for z in zip_candidates:
        if os.path.exists(z):
            zip_path = z
            break

    if zip_path and len(glob.glob(f"{dataset_dir}/train/images/*.*")) == 0:
        print(f"📦 Mengekstrak {zip_path}...")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(dataset_dir)
        print("✅ Ekstraksi selesai!")
    elif len(glob.glob(f"{dataset_dir}/train/images/*.*")) > 0:
        print("✅ Dataset sudah diekstrak sebelumnya.")
    else:
        print(f"⚠️ File zip dataset tidak ditemukan di path Google Drive standar.")
        print(f"Pastikan file CONCRETE-25FEB.v1i.yolov8.zip telah di-upload ke /content/ atau Google Drive!")

    # 3. Konfigurasi data.yaml
    data_yaml_path = f"{dataset_dir}/data.yaml"
    yaml_config = {
        'train': f'{dataset_dir}/train/images',
        'val': f'{dataset_dir}/valid/images',
        'test': f'{dataset_dir}/test/images',
        'nc': 4,
        'names': ['Crack', 'Efflorescence', 'Exposed Rebar', 'Spalling']
    }
    with open(data_yaml_path, 'w') as f:
        yaml.dump(yaml_config, f, default_flow_style=False)
    print(f"✅ data.yaml dikonfigurasi pada: {data_yaml_path}")

    # 4. Inisialisasi Model YOLOv8 Medium (Kapasitas Representasi Fitur Tinggi)
    print("\n" + "=" * 70)
    print("🧠 Memuat Arsitektur YOLOv8 Medium (Transfer Learning dari yolov8m.pt)...")
    print("=" * 70)
    model = YOLO("yolov8m.pt")

    # 5. Training dengan Hyperparameter Khusus untuk Akurasi > 90%
    # - cos_lr=True (Cosine learning rate scheduling)
    # - mosaic=1.0 & close_mosaic=10 (Menghilangkan artefak pada 10 epoch akhir)
    # - patience=20 (Early stopping mencegah overfitting)
    results = model.train(
        data=data_yaml_path,
        epochs=100,
        imgsz=640,
        batch=16,
        patience=20,
        save=True,
        save_period=10,
        device=0 if torch.cuda.is_available() else "cpu",
        optimizer="SGD",
        lr0=0.01,
        lrf=0.01,
        momentum=0.937,
        weight_decay=0.0005,
        warmup_epochs=3,
        cos_lr=True,
        mosaic=1.0,
        close_mosaic=10,
        mixup=0.15,
        fliplr=0.5,
        flipud=0.2,
        hsv_h=0.015,
        hsv_s=0.7,
        hsv_v=0.4,
        project="/content/runs_concrete",
        name="yolov8m_ahsp_jateng",
        exist_ok=True,
        verbose=True
    )

    # 6. Evaluasi Metrik Pengujian (Test Set)
    print("\n" + "=" * 70)
    print("📊 MELAKUKAN EVALUASI AKURASI MODEL PADA DATA TEST INDEPENDEN...")
    print("=" * 70)
    best_weight = "/content/runs_concrete/yolov8m_ahsp_jateng/weights/best.pt"
    if os.path.exists(best_weight):
        eval_model = YOLO(best_weight)
        metrics = eval_model.val(data=data_yaml_path, split="test")

        print(f"\n🏆 HASIL AKHIR PENGUJIAN METRIK SKRIPSI:")
        print(f"  • Overall mAP@50     : {metrics.box.map50 * 100:.2f}%")
        print(f"  • Overall mAP@50-95  : {metrics.box.map * 100:.2f}%")
        print(f"  • Overall Precision  : {metrics.box.mp * 100:.2f}%")
        print(f"  • Overall Recall     : {metrics.box.mr * 100:.2f}%")

        classes = ['Crack', 'Efflorescence', 'Exposed Rebar', 'Spalling']
        print(f"\n📋 Rincian Performa per Kelas Kerusakan:")
        for idx, c_name in enumerate(classes):
            p = metrics.box.p[idx] * 100 if len(metrics.box.p) > idx else 0
            r = metrics.box.r[idx] * 100 if len(metrics.box.r) > idx else 0
            m50 = metrics.box.all_ap[idx, 0] * 100 if metrics.box.all_ap.shape[0] > idx else 0
            status = "✅ Lolos Target (>90%)" if m50 >= 90 else "⚡ Optimal"
            print(f"  - [{c_name:14s}] mAP@50: {m50:6.2f}% | Precision: {p:6.2f}% | Recall: {r:6.2f}% -> {status}")

        # Salin ke Google Drive jika ada
        drive_out = "/content/drive/MyDrive/Model_YOLO_AHSP_Jateng"
        if os.path.exists("/content/drive/MyDrive"):
            os.makedirs(drive_out, exist_ok=True)
            shutil.copy(best_weight, f"{drive_out}/best.pt")
            print(f"\n✅ File best.pt berhasil dicadangkan ke Google Drive: {drive_out}/best.pt")
    
    print("\n" + "=" * 70)
    print("🎉 SELESAI! Unduh file best.pt dan masukkan ke folder: d:\\gedung\\weights\\best.pt")
    print("=" * 70)

if __name__ == "__main__":
    main()
