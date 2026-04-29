import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("Cleaning Laporan Laba Rugi")

st.write("*Realisasi*: data laporan laba rugi dari modul Accounting. File yang diupload harus dalam format Excel (.xlsx) dengan struktur: **:green[perusahaan_periode]**, contoh: **:green[bimp_des25.xlsx]**")

uploaded_files = st.file_uploader("Upload File Laporan Laba Rugi (Excel)", type=['xlsx'], accept_multiple_files=True)

if not uploaded_files:
    st.warning("Silakan upload minimal 1 file")

if uploaded_files:
    df_list = []
    periode_list = []
    
    pd.set_option("display.max_columns", None)

    with st.spinner("Sedang memproses..."):
        for uploaded_file in uploaded_files:
            # 1) Buka file Excel
            df = pd.read_excel(uploaded_file, skiprows=1)

            # 2) Ambil Tahun dan Bulan dari nama file
            file_name = uploaded_file.name  # contoh: bimp_des25.xlsx
            file_name = file_name.replace(".xlsx", "").lower()

            parts = file_name.split("_")  # ["bimp", "des25"]

            bulan_map = {
                "jan": "Januari", "feb": "Februari", "mar": "Maret",
                "apr": "April", "mei": "Mei", "jun": "Juni",
                "jul": "Juli", "agu": "Agustus", "sep": "September",
                "okt": "Oktober", "nov": "November", "des": "Desember"
            }

            bulan = None
            tahun = None

            # ambil bagian terakhir (des25)
            last = parts[-1]
            periode_list.append(last)

            match = re.match(r"([a-z]{3})(\d{2})", last)
            if match:
                bulan_code = match.group(1)
                tahun_code = match.group(2)

                bulan = bulan_map.get(bulan_code[:3])  # ambil 3 huruf awal
                tahun = int("20" + tahun_code) if len(tahun_code) == 2 else int(tahun_code)

            if bulan is None or tahun is None:
                st.warning(f"Gagal ekstrak bulan/tahun dari nama file: {file_name}. Pastikan format nama file benar.")
                st.stop()

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

            df = df[df["Overview"] != ""]

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
        
        # Kondisi untuk filename berdasarkan jumlah periode yang diproses
        if len(periode_list) == 1:
            periode_str = periode_list[0]
        else:
            periode_str = f"{periode_list[0]}-{periode_list[-1]}"

        st.download_button(
            label="Download Hasil Cleaning",
            data=output.getvalue(),
            file_name=f"{parts[0]}_{periode_str}_cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cleaned_realisasi",
            icon=":material/download:",
            type="primary"
        )