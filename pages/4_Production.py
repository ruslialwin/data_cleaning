import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("Cleaning Produksi Harian")

st.write("*Production*: data laporan produksi harian dari modul Pabrik. File yang diupload harus dalam format Excel (.xlsx) dengan struktur: **:green[produksi_perusahaan_periode]**, contoh: **:green[produksi_bimp_des25.xlsx]**")

uploaded_file = st.file_uploader("Upload File Produksi Harian (Excel)", type=['xlsx'])

if not uploaded_file:
    st.warning("Silakan upload file")

if uploaded_file: 
    pd.set_option("display.max_columns", None)

    with st.spinner("Sedang memproses..."):
        # 1) Membaca file excel dengan melewati 2 baris pertama
        df = pd.read_excel(uploaded_file, skiprows=2)

        file_name = uploaded_file.name  # contoh: produksi_bimp_des25.xlsx
        file_name = file_name.replace(".xlsx", "_cleaned.xlsx")  # hasil: produksi_bimp_des25_cleaned.xlsx

        # 2) Mengganti nama kolom pertama menjadi "Tanggal"
        df.rename(columns={df.columns[0]: "Tanggal"}, inplace=True)

        # 3) Menghapus baris yang memiliki nilai "JUMLAH" pada kolom "Tanggal"
        df = df[df["Tanggal"] != "JUMLAH"]

        # 4) Mengganti nama bulan dalam kolom "Tanggal" dari bahasa Indonesia ke bahasa Inggris
        bulan_map = {
            'Jan': 'January', 'Feb': 'February', 'Mar': 'March',
            'Apr': 'April', 'Mei': 'May', 'Jun': 'June',
            'Jul': 'July', 'Agt': 'August', 'Sep': 'September',
            'Okt': 'October', 'Nov': 'November', 'Des': 'December'
        }

        df["Tanggal"] = df["Tanggal"].replace(bulan_map, regex=True).str.strip()
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], format="%d %B %Y")
        df["Tanggal"] = df["Tanggal"].dt.strftime("%d/%m/%Y")
        
        # 5) Menambahkan kolom "Tahun", "Bulan", dan "Kapasitas Olah Maksimal"
        bulan_map = {
            1: 'Januari', 2: 'Februari', 3: 'Maret',
            4: 'April', 5: 'Mei', 6: 'Juni',
            7: 'Juli', 8: 'Agustus', 9: 'September',
            10: 'Oktober', 11: 'November', 12: 'Desember'
        }

        df.insert(1, "Tahun", pd.to_datetime(df["Tanggal"], format="%d/%m/%Y").dt.year)
        df.insert(2, "Bulan", pd.to_datetime(df["Tanggal"], format="%d/%m/%Y").dt.month.map(bulan_map)) 
        df["Kapasitas Olah Maksimal"] = 60000 

        # 6) Mengganti koma dengan titik dan melakukan pembulatan 2 angka di belakang koma pada kolom "Total Jam Operasi"
        df["Total Jam Operasi"] = (
            df["Total Jam Operasi"]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .astype(float)
            .round(2)
            .astype(str)
            .str.rstrip('0')
            .str.rstrip('.')
        )

    st.success("Selesai!")
    st.dataframe(df.head())

    # Tombol Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Production")

        workbook  = writer.book
        worksheet = writer.sheets["Production"]

        format_angka = workbook.add_format({
            "num_format": '#,##0'
        })
        
        columns = ["TBS diolah", "TBS olah After Grading", "Produksi CPO", "Produksi Kernel", "Kapasitas Olah Maksimal"]
        for column in columns:
            col_idx = df.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, 18, format_angka)

    st.download_button(
        label="Download Hasil Cleaning",
        data=output.getvalue(),
        file_name=f"{file_name}", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_cleaned_production",
        icon=":material/download:",
        type="primary"
    )