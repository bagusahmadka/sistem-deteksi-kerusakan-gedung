PETUNJUK PENEMPATAN MODEL YOLOV8 HASIL TRAINING:
==================================================
Setelah Anda selesai melakukan training di Google Colab menggunakan script/notebook:
Train_YOLOv8_Concrete_Defects_Colab.ipynb

Unduh file "best.pt" dari Google Colab, lalu letakkan di folder ini dengan nama:
d:\gedung\weights\best.pt

Aplikasi website Streamlit (app.py) akan secara otomatis mendeteksi model ini dan
menggunakannya untuk inferensi deteksi visual real-time (bounding box, confidence score,
serta klasifikasi kerusakan sesuai AHSP Jawa Tengah 2023).
