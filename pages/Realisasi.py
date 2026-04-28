import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.title("Cleaning Laporan Laba Rugi")

# st.markdown(
#     ":blue-badge[:green[**Realisasi**]] adalah fitur untuk membersihkan data laporan laba rugi yang diupload oleh user. Fitur ini akan memproses file Excel yang diupload, mengekstrak informasi tahun dan bulan, serta mengkategorikan data berdasarkan aturan tertentu. Setelah proses selesai, user dapat mendownload hasilnya dalam format Excel dengan nama file yang sudah ditentukan sesuai dengan pilihan perusahaan, bulan, dan tahun."
# )

col1, col2, col3 = st.columns(3)
    
with col1:
    perusahaan = st.selectbox(
        "Pilih perusahaan",
        ["BIMP", "BIMS", "BIMR", "KPNJ", "MUL"],
        index=0
    )

with col2:
    bulan = st.selectbox(
        "Pilih periode bulan",
        ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    )

with col3:
    tahun = st.selectbox(
        "Pilih periode tahun",
        ["2025", "2026"]
    )

bulan_map = {
    "Januari": "jan",
    "Februari": "feb",
    "Maret": "mar",
    "April": "apr",
    "Mei": "mei",
    "Juni": "jun",
    "Juli": "jul",
    "Agustus": "agu",
    "September": "sep",
    "Oktober": "okt",
    "November": "nov",
    "Desember": "des"
}

tahun_short = tahun[-2:]

file_name = f"{perusahaan.lower()}_{bulan_map[bulan]}{tahun_short}_cleaned.xlsx"

st.markdown(
    f"Nama file: :blue-badge[{file_name}]"
)

st.divider()

uploaded_files = st.file_uploader("Upload File Laporan Laba Rugi (Excel)", type=['xlsx'], accept_multiple_files=True)

if not uploaded_files:
    st.warning("Silakan upload minimal 1 file")

if uploaded_files:
    df_list = []
    
    with st.spinner("Sedang memproses..."):
        for uploaded_file in uploaded_files:
            # 1) Buka file Excel
            df_tanggal = pd.read_excel(
                uploaded_file,
                header=None,
                nrows=1
            )

            df = pd.read_excel(uploaded_file, skiprows=1)

            # 2) Ambil Tahun dan Bulan dari df_tanggal
            text = df_tanggal.iloc[0, 1]
            
            parts = text.split()

            bulan_arr = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
            tahun_arr = ["2025", "2026"]
            
            bulan = next((p.capitalize() for p in parts if p.capitalize() in bulan_arr), None)
            tahun = next((int(p) for p in parts if p.isdigit() and p in tahun_arr), None)

            # 3) Simpan baris yang tidak diawali dengan huruf
            df = df[~df[df.columns[0]].astype(str).str.match(r'^[^\d]+')]

            # 4) Tambahkan kolom baru
            df.insert(0, "Tahun", tahun)
            df.insert(1, "Bulan", bulan)
            df.insert(2, "Overview", "")
            df.insert(3, "Deskripsi", "")
            df.insert(4, "No. Akun", "")
            df.columns.values[5] = "Rincian Deskripsi"
            df.columns.values[6] = "Realisasi Biaya"

            # 5) Pisah kolom "Rincian Deskripsi" menjadi "No. Akun" dan "Rincian Deskripsi"
            df["No. Akun"] = df["Rincian Deskripsi"].str.extract(r'^([\d\.]+)')
            df["Rincian Deskripsi"] = df["Rincian Deskripsi"].str.replace(r'^([\d\.]+)\s*', '', regex=True)
            
            # 6) Kategorikan "Deskripsi" berdasarkan aturan yang diberikan
            df.loc[df["No. Akun"].str.startswith("41"), "Deskripsi"] = "Pendapatan"

            df.loc[df["Rincian Deskripsi"].str.contains("Pembelian", case=False, na=False), "Deskripsi"] = "Pembelian"
            df.loc[df["Rincian Deskripsi"].str.contains("Harga Pokok", case=False, na=False), "Deskripsi"] = "Persediaan Awal"
            df.loc[df["Rincian Deskripsi"].str.contains("Transit", case=False, na=False), "Deskripsi"] = "Persediaan Akhir"
            df.loc[df["No. Akun"].str.startswith("51.02"), "Deskripsi"] = "Biaya Pembelian"
            df.loc[df["No. Akun"].str.startswith("51.03"), "Deskripsi"] = "Biaya Gaji"
            df.loc[df["No. Akun"].str.startswith("52"), "Deskripsi"] = "Biaya Overhead"

            df.loc[df["No. Akun"].str.startswith("62"), "Deskripsi"] = "Beban Penjualan"

            df.loc[df["No. Akun"].str.startswith("70"), "Deskripsi"] = "Beban Administrasi"

            df.loc[df["No. Akun"].str.startswith("80"), "Deskripsi"] = "Pendapatan Lain-lain"
            df.loc[df["No. Akun"].eq("41.01.01.0000.11"), "Deskripsi"] = "Pendapatan Lain-lain"

            df.loc[df["No. Akun"].str.startswith("81"), "Deskripsi"] = "Beban Lain-lain"
            df.loc[df["No. Akun"].eq("81.07.00.0000.02"), "Deskripsi"] = "Pendapatan Lain-lain"

            # 7) Kategorikan "Overview" berdasarkan "Deskripsi"
            df["Deskripsi"].astype(str)

            mapping = {
                "Pendapatan/Penjualan": ["Pendapatan"],
                "Beban Pokok Pendapatan/COGS": ["Pembelian", "Persediaan Awal", "Persediaan Akhir", "Biaya Pembelian", "Biaya Gaji", "Biaya Overhead"],
                "Beban Operasional": ["Beban Penjualan", "Beban Administrasi"],
                "Pendapatan Non-Operasional": ["Pendapatan Lain-lain"],
                "Beban Non-Operasional": ["Beban Lain-lain"]
            }

            for overview, deskripsi in mapping.items():
                pattern = "|".join(deskripsi)
                df.loc[df["Deskripsi"].str.contains(pattern, case=False, na=False), "Overview"] = overview

            # Hapus baris yang memiliki nilai kosong 
            df = df.dropna()

            df_list.append(df)

        # gabungkan semua
        df_final_all = pd.concat(df_list, ignore_index=True)

        st.success("Selesai!")
        st.dataframe(df_final_all.head())

        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final_all.to_excel(writer, index=False, sheet_name="Realisasi")

            workbook  = writer.book
            worksheet = writer.sheets["Realisasi"]

            format_angka = workbook.add_format({
                "num_format": '#,##0.00;(#,##0.00);-'
            })

            col_idx = df_final_all.columns.get_loc("Realisasi Biaya")
            worksheet.set_column(col_idx, col_idx, 18, format_angka)
        
        st.download_button(
            label="Download Hasil Cleaning",
            data=output.getvalue(),
            file_name=f"{file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cleaned_realisasi"
        )