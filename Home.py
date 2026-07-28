import streamlit as st
from auth import get_authenticator

st.set_page_config(
    page_title="Odoo Data Cleaner",
    layout="wide"
)

authenticator = get_authenticator()

authenticator.login(location="main")

if st.session_state.get("authentication_status"):

    authenticator.logout(location="sidebar")

    st.title("Odoo Data Cleaning Tools")

    st.markdown("""
    Tools otomatisasi cleaning data Excel Odoo.
    Pilih jenis data yang ingin dibersihkan melalui sidebar di sebelah kiri.

    **Menu yang tersedia:**
    1. **Realisasi**: Cleaning data laporan laba dan rugi dari modul Accounting.
    2. **Processing & Labour**: Data realisasi hanya biaya gaji, overhead dan beban administrasi (exclude biaya penyusutan) dengan penambahan kolom tipe.
    3. **Realisasi Akun Analitik**: Cleaning data artikel jurnal dari modul Accounting.
    4. **Production**: Data laporan produksi harian dari modul Pabrik.
    5. **FFB Price**: Mengambil data harga TBS dari modul Odoo.
    """)

elif st.session_state.get("authentication_status") is False:
    st.error("Username/password incorrect")