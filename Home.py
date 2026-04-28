import streamlit as st

st.set_page_config(page_title="Odoo Data Cleaner", layout="wide")

st.title("Odoo Data Cleaning Tools")
st.markdown("""
Tools otomatisasi cleaning data Excel Odoo. 
Pilih jenis data yang ingin dibersihkan melalui sidebar di sebelah kiri.

**Menu yang tersedia:**
1. **Realisasi**: Cleaning data laporan laba dan rugi dari modul Accounting.
2. **Processing & Labour**: Data realisasi hanya biaya gaji, overhead dan beban administrasi (exclude biaya penyusutan) dengan penambahan kolom tipe.
3. **Realisasi Akun Analitik**: Cleaning data artikel jurnal dari modul Accounting.
""")