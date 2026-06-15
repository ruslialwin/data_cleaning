import streamlit as st
import pandas as pd
import requests
import re
from io import BytesIO
from datetime import datetime

st.title("Cleaning Laporan Laba Rugi")

source = st.radio(
    "Sumber Data",
    ["Upload File Excel", "API"],
    horizontal=True
)

if source == "Upload File Excel":
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

elif source == "API":
    st.info(
        "Data laporan laba rugi diambil langsung dari API Dashboard. "
        "Klik **Ambil Data Laba Rugi** untuk memuat data dan mengunduh hasil dalam format Excel (.xlsx). "
        "Nama file akan dibuat otomatis berdasarkan perusahaan dan periode, misalnya "
        "`bimp_jul25.xlsx`."
    )

    company_options = {
    "BIMP": "https://dashboard.mahkotagroup.com/api/dashboard/realisasi-bimp?mode=live",
    "BIMR": "https://dashboard.mahkotagroup.com/api/dashboard/realisasi-bimr?mode=live",
    "BIMS": "https://dashboard.mahkotagroup.com/api/dashboard/realisasi-bims?mode=live",
    "MUL": "https://dashboard.mahkotagroup.com/api/dashboard/realisasi-mul?mode=live",
    "KPNJ": "https://dashboard.mahkotagroup.com/api/dashboard/realisasi-kpnj?mode=live"
    }

    selected_company = st.selectbox(
        "Pilih Perusahaan",
        list(company_options.keys())
    )

    API_URL = company_options[selected_company]

    if st.button("Ambil Data Laba Rugi"):
        with st.spinner("Mengambil data dari API...."):
            response = requests.get(API_URL)

            if response.status_code != 200:
                st.error(f"Gagal mengambil data. Status: {response.status_code}")
                st.stop()

            result = response.json()

            # langsung pakai hasil transformasi dari API
            df = pd.DataFrame(result["data"])

            st.write(df.columns.tolist())

            # Balik tanda nominal untuk Pendapatan dan Pendapatan Lain-lain
            mask_pendapatan = df["Deskripsi"].isin(["Pendapatan", "Pendapatan Lain-lain"])
            df.loc[mask_pendapatan, "Realisasi Biaya"] = df.loc[mask_pendapatan, "Realisasi Biaya"] * -1

            # Susun kolom akhir
            final_cols = [
                "Tahun", "Bulan", "Overview", "Deskripsi", "No. Akun", "Rincian Deskripsi", "Realisasi Biaya"
            ]


            df = df[final_cols].copy()


            # mengambil informasi id perusahaan dari respons API
            company_id = result["config"]["context"]["allowed_company_ids"][0]

            company_map = {
                2: "mul",
                4: "bimp",
                5: "bimr",
                6: "bims",
                10: "kpnj"
            }

            # mengambil informasi periode dari respons API
            custom_domain = result["config"]["customDomain"]

            start_date = None
            end_date = None

            for item in custom_domain:
                if isinstance(item, list):
                    if item[0] == "date" and item[1] == ">=":
                        start_date = item[2]

                    if item[0] == "date" and item[1] == "<=":
                        end_date = item[2]

            start_dt = datetime.strptime(start_date, "%Y-%m-%d")
            end_dt = datetime.strptime(end_date, "%Y-%m-%d")
            
            if start_dt.strftime("%b").lower() == end_dt.strftime("%b").lower():
                periode = start_dt.strftime("%b%y").lower()
            else:
                periode = start_dt.strftime("%b%y").lower() + "-" + end_dt.strftime("%b%y").lower()

            filename=f"{company_map.get(company_id, company_id)}_{periode}_terekam_cleaned.xlsx"

            st.caption(
                f"Perusahaan: {company_map.get(company_id, company_id).upper()} | "
                f"Periode: {start_date} s.d. {end_date}"
            )

        st.success("Selesai!")
        st.dataframe(df.head())

        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Realisasi")

            workbook  = writer.book
            worksheet = writer.sheets["Realisasi"]

            format_angka = workbook.add_format({
                "num_format": '#,##0.00;(#,##0.00);-'
            })
            
            col_idx = df.columns.get_loc("Realisasi Biaya")
            worksheet.set_column(col_idx, col_idx, 18, format_angka)

        st.write("Company ID :", company_id)
        st.write("Start Date :", start_date)
        st.write("End Date :", end_date)

        st.download_button(
            label=f"Download Hasil Cleaning: {filename}",
            data=output.getvalue(),
            file_name=filename, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cleaned_realisasi_terekam",
            icon=":material/download:",
            type="primary"
        )