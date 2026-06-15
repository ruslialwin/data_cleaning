import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime

st.title("Cleaning Produksi Harian")

source = st.radio(
    "Sumber Data",
    ["Upload File Excel", "API"],
    horizontal=True
)

if source == "Upload File Excel":
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

elif source == "API":
    st.info(
        "Data produksi harian diambil langsung dari API Dashboard. "
        "Klik **Ambil Data Produksi** untuk memuat data dan mengunduh hasil dalam format Excel (.xlsx). "
        "Nama file akan dibuat otomatis berdasarkan perusahaan dan periode, misalnya "
        "`produksi_bimp_jul25.xlsx`."
    )

    company_options = {
    "BIMP": "https://dashboard.mahkotagroup.com/api/dashboard/production-bimp?mode=live",
    "BIMR": "https://dashboard.mahkotagroup.com/api/dashboard/production-bimr?mode=live",
    "BIMS": "https://dashboard.mahkotagroup.com/api/dashboard/production-bims?mode=live",
    "MUL": "https://dashboard.mahkotagroup.com/api/dashboard/production-mul?mode=live",
    "KPNJ": "https://dashboard.mahkotagroup.com/api/dashboard/production-kpnj?mode=live"
    }

    selected_company = st.selectbox(
        "Pilih Perusahaan",
        list(company_options.keys())
    )

    API_URL = company_options[selected_company]

    if st.button("Ambil Data Produksi"):
        with st.spinner("Mengambil data dari API...."):
            response = requests.get(API_URL)

            if response.status_code != 200:
                st.error(f"Gagal mengambil data. Status: {response.status_code}")
                st.stop()

            result = response.json()

            # langsung pakai hasil transformasi dari API
            df = pd.DataFrame(result["data"])

            st.write(df.columns.tolist())

            # format tanggal
            df["Tanggal"] = pd.to_datetime(df["Tanggal"])
            df["Tanggal"] = df["Tanggal"].dt.strftime("%d/%m/%Y")

            # format Total Jam Operasi
            df["Total Jam Operasi"] = (
                df["Total Jam Operasi"]
                .astype(float)
                .round(2)
                .astype(str)
                .str.rstrip("0")
                .str.rstrip(".")
            )

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
                    if item[0] == "production_date" and item[1] == ">=":
                        start_date = item[2]

                    if item[0] == "production_date" and item[1] == "<=":
                        end_date = item[2]

            start_dt = datetime.strptime(start_date, "%Y-%m-%d")
            end_dt = datetime.strptime(end_date, "%Y-%m-%d")

            periode = start_dt.strftime("%d %b %y").lower() + " - " + end_dt.strftime("%d %b %y").lower()

            filename=f"produksi_{company_map.get(company_id, company_id)}_{periode}_cleaned.xlsx"

            st.caption(
                f"Perusahaan: {company_map.get(company_id, company_id).upper()} | "
                f"Periode: {start_date} s.d. {end_date}"
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
            
            columns = ["TBS Diolah", "TBS olah After Grading", "Produksi CPO", "Produksi Kernel", "Kapasitas Olah Maksimal"]
            for column in columns:
                col_idx = df.columns.get_loc(column)
                worksheet.set_column(col_idx, col_idx, 18, format_angka)

        st.write("Company ID :", company_id)
        st.write("Start Date :", start_date)
        st.write("End Date :", end_date)

        st.download_button(
            label=f"Download Hasil Cleaning: {filename}",
            data=output.getvalue(),
            file_name=filename, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cleaned_production",
            icon=":material/download:",
            type="primary"
        )