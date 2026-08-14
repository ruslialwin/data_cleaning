# Odoo Data Cleaning Tools 
### Tools otomatisasi cleaning data Excel Odoo. 
### Menu Halaman:
1. **Realisasi**: Cleaning data laporan laba dan rugi dari modul Accounting.
2. **Processing & Labour**: Data realisasi hanya biaya gaji, overhead dan beban administrasi (exclude biaya penyusutan) dengan penambahan kolom tipe.
3. **Realisasi Akun Analitik**: Cleaning data artikel jurnal dari modul Accounting.
4. **Production**: Data laporan produksi harian dari modul Pabrik.

## 🛠️ Dibuat Dengan
- **Streamlit**: Untuk membangun antarmuka web interaktif berbasis Python.
- **Python**: Untuk pemrosesan data dan logika aplikasi.
- **Pandas**: Untuk cleaning, transformasi, dan analisis data Excel Odoo.
- **OpenPyXL / XlsxWriter**: Untuk membaca, memproses, dan export file Excel.
- **NumPy**: Untuk operasi numerik dan manipulasi data.
- **Docker**: Untuk containerization dan menjalankan aplikasi dalam environment yang konsisten.
  
## 🚀 Menjalankan Aplikasi

### 🌐 Streamlit Cloud

Aplikasi dapat langsung diakses melalui:
[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://odoo-data-cleaning.streamlit.app/)

### 🐳 Docker

Pastikan **Docker** sudah terinstall di komputer.

#### 1. Clone repository

```bash
git clone https://github.com/ruslialwin/odoo-data-cleaning.git
cd odoo-data-cleaning
```

#### 2. Docker Build

```bash
docker compose up --build
```
#### 3. Akses Aplikasi

```bash
http://localhost:8501
```
