import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.title("Cleaning Processing & Labour") 


st.write("*Processing & Labour* : data realisasi hanya biaya gaji, overhead dan beban administrasi (exclude biaya penyusutan) dengan penambahan kolom tipe. File yang diupload harus **file realisasi yang telah di-cleaned** dengan struktur: **:green[perusahaan_periode_cleaned]**, contoh: **:green[bimp_des25_cleaned.xlsx]**")

uploaded_files = st.file_uploader("Upload File Realisasi yang telah di-cleaned (Excel)", type=['xlsx'], accept_multiple_files=True)

if not uploaded_files:
    st.warning("Silakan upload minimal 1 file")

if uploaded_files:
    df_list = []

    pd.set_option("display.max_columns", None)

    with st.spinner("Sedang memproses..."):
        for uploaded_file in uploaded_files:
            # 1) Buka file Excel
            df_realisasi = pd.read_excel(uploaded_file)

            # 2) Mengambil nama file untuk digunakan sebagai nama file hasil download
            file_name = uploaded_file.name  # contoh: bimp_des25_cleaned.xlsx

            # tambah prolab_ di depan nama file dan hapus _cleaned
            file_name = "prolab_" + file_name.replace("_cleaned.xlsx", "").replace(".xlsx", "")

            # 2) Mengambil data yang tidak mengandung kata "Peny" pada kolom "Rincian Deskripsi"
            df_realisasi = df_realisasi[~df_realisasi["Rincian Deskripsi"].str.contains("Peny", case=False, na=False)]

            # 3) Mengambil data yang memiliki nilai "Deskripsi" sesuai dengan daftar yang diberikan
            df_filtered = df_realisasi[
                df_realisasi["Deskripsi"].isin([
                    "Biaya Gaji",
                    "Biaya Overhead",
                    "Beban Administrasi"
                ])
            ]

            # 4) Menambahkan kolom "Tipe" berdasarkan nilai pada kolom "Deskripsi"
            mapping = {
                "Biaya Gaji": "Labour Cost",
                "Beban Administrasi": "Labour Cost",
                "Biaya Overhead": "Processing Cost"
            }

            # Buat kolom baru dulu
            df_filtered["Tipe_temp"] = df_filtered["Deskripsi"].map(mapping)

            # Ambil posisi kolom "Deskripsi"
            pos = df_filtered.columns.get_loc("Deskripsi")

            # Insert kolom "Tipe" di sebelahnya
            df_filtered.insert(pos + 1, "Tipe", df_filtered.pop("Tipe_temp"))

            df_list.append(df_filtered)

        # gabungkan semua
        df_final_all = pd.concat(df_list, ignore_index=True)

        st.success("Selesai!")
        st.dataframe(df_final_all.head())

        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final_all.to_excel(writer, index=False, sheet_name="Processing & Labour")

            workbook  = writer.book
            worksheet = writer.sheets["Processing & Labour"]

            format_angka = workbook.add_format({
                "num_format": '#,##0.00;(#,##0.00);-'
            })

            col_idx = df_final_all.columns.get_loc("Realisasi Biaya")
            worksheet.set_column(col_idx, col_idx, 18, format_angka)

        st.download_button(
            label=f"Download {file_name}.xlsx",
            data=output.getvalue(),
            file_name=f"{file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_processing_labour",
            icon=":material/download:",
            type="primary"
        )