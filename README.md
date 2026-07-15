# 🏢 Executive Dashboard: Analisis Mesin Diesel WPO (Waste Plastic Oil)

Dashboard interaktif berbasis **Streamlit** untuk memvisualisasikan, mengeksplorasi, dan menganalisis dataset termodinamika performa serta emisi mesin diesel yang menggunakan campuran *Waste Plastic Oil (WPO) / Plastic Pyrolysis Oil (PPO)*. 

Data yang digunakan dalam dashboard ini telah melalui proses **Augmentasi Polinomial Orde-2** yang ketat untuk mempertahankan hukum fisika dan termodinamika mesin.

## 📂 Struktur Folder Proyek
Pastikan file Anda tersusun di dalam satu folder dengan struktur sebagai berikut:

```text
📁 nama-folder-proyek/
 ├── 📄 app.py                                  # Script utama Streamlit
 ├── 📄 requirements.txt                        # Daftar library Python yang dibutuhkan
 ├── 📄 Polinomial_Augmented_Dataset.xlsx       # Dataset hasil augmentasi (Wajib ada)
 └── 📄 README.md                               # File panduan ini
```

## ⚙️ Persyaratan Sistem (Prerequisites)
Sebelum menjalankan aplikasi, pastikan komputer Anda telah terinstal:
* **Python** (Versi 3.9 atau lebih baru)
* **pip** (Python Package Installer)

---

## 🚀 Cara Instalasi & Menjalankan Offline (Localhost)

Ikuti langkah-langkah berikut melalui Terminal (Mac/Linux) atau Command Prompt / PowerShell (Windows):

### 1. Buka Terminal dan Arahkan ke Folder Proyek
Buka terminal dan gunakan perintah `cd` untuk masuk ke folder tempat file-file di atas disimpan.
```bash
cd path/to/nama-folder-proyek
```

### 2. Buat Virtual Environment (Sangat Direkomendasikan)
Praktik terbaik dalam *software engineering* adalah memisahkan instalasi *library* proyek ini dari sistem utama Anda.
* **Pengguna Windows:**
  ```bash
  python -m venv venv
  venv\Scriptsctivate
  ```
* **Pengguna Mac/Linux:**
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  ```
*(Catatan: Jika virtual environment aktif, akan muncul tulisan `(venv)` di sebelah kiri prompt terminal Anda).*

### 3. Install Dependencies (Library)
Install semua library yang terdaftar di file `requirements.txt` dengan perintah:
```bash
pip install -r requirements.txt
```
*Tunggu hingga proses instalasi (Pandas, Streamlit, Plotly, dll.) selesai 100%.*

### 4. Jalankan Aplikasi Streamlit
Ketik perintah berikut untuk meluncurkan dashboard:
```bash
streamlit run app.py
```

### 5. Buka di Browser
Setelah perintah dijalankan, Streamlit akan secara otomatis membuka jendela browser baru dengan alamat `http://localhost:8501`. 
Jika tidak terbuka otomatis, Anda bisa mengklik atau men-copy URL *Local URL* yang muncul di terminal ke Google Chrome atau browser pilihan Anda.

---

## 🛑 Cara Menghentikan Aplikasi
Untuk mematikan server lokal Streamlit, kembali ke Terminal/Command Prompt Anda, lalu tekan kombinasi *keyboard*:
**`Ctrl + C`**

---

## 💡 Pemecahan Masalah Umum (Troubleshooting)

* **Error `FileNotFoundError: [Errno 2] No such file or directory: 'Polinomial_Augmented_Dataset.xlsx'`**
  * **Solusi**: Pastikan nama file excel Anda sama persis (huruf besar/kecil berpengaruh) dan file tersebut berada di dalam folder yang sama dengan `app.py`.
* **Error `ModuleNotFoundError: No module named '...'`**
  * **Solusi**: Pastikan Anda sudah menjalankan perintah `pip install -r requirements.txt`. Jika menggunakan virtual environment, pastikan environment tersebut dalam keadaan aktif saat instalasi dan saat menjalankan aplikasi.
* **Port 8501 sudah terpakai (Port is already in use)**
  * **Solusi**: Streamlit akan otomatis mencari port kosong (misal 8502). Cek URL yang tertera di terminal Anda.
