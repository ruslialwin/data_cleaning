import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime

st.title("Cleaning Processing & Labour") 

source = st.radio(
    "Sumber Data",
    ["Upload File Excel", "API"],
    horizontal=True
)

if source == "Upload File Excel":
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

                # 3) Mengambil data yang tidak mengandung kata "Peny" pada kolom "Rincian Deskripsi"
                df_realisasi = df_realisasi[~df_realisasi["Rincian Deskripsi"].str.contains("Peny", case=False, na=False)]

                # 4) Mengambil data yang memiliki nilai "Deskripsi" sesuai dengan daftar yang diberikan
                df_filtered = df_realisasi[
                    df_realisasi["Deskripsi"].isin([
                        "Biaya Gaji",
                        "Biaya Overhead",
                        "Beban Administrasi"
                    ])
                ]

                # 5) Menambahkan kolom "Tipe" berdasarkan nilai pada kolom "Deskripsi"
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

elif source == "API":
    st.info(
        "Data processing & labour diambil langsung dari API Dashboard. "
        "Klik **Ambil Data Processing & Labour** untuk memuat data dan mengunduh hasil dalam format Excel (.xlsx). "
        "Nama file akan dibuat otomatis berdasarkan perusahaan dan periode, misalnya "
        "`prolab_bimp_jul25.xlsx`."
    )

    company_options = {
        "BIMP": [
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bimp-20250701-20251231",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bimp-20260101-20260630",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bimp-20260701-20261231"
        ],
        "BIMR": [
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bimr-20250701-20251231",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bimr-20260101-20260630",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bimr-20260701-20261231"
        ],
        "BIMS": [
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bims-20250701-20250930",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bims-20251001-20251231",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bims-20260101-20260331",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bims-20260401-20260630",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bims-20260701-20260930",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-bims-20261001-20261231"
        ],
        "MUL": [
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-mul-20250701-20251231",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-mul-20260101-20260630",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-mul-20260701-20261231"
        ],
        "KPNJ": [
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-kpnj-20250701-20251231",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-kpnj-20260101-20260331",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-kpnj-20260401-20260630",
            "https://dashboard.mahkotagroup.com/api/powerbi-feed/processing-labour-kpnj-20260701-20261231"
        ]
    }

    selected_company = st.selectbox(
        "Pilih Perusahaan",
        list(company_options.keys())
    )

    API_URLS = company_options[selected_company]

    if st.button("Ambil Data Processing & Labour"):
        with st.spinner("Mengambil data dari API...."):
            df_list = []
            start_dates = []
            end_dates = []
            
            for API_URL in API_URLS:
                response = requests.get(API_URL)

                if response.status_code != 200:
                    st.error(f"Gagal mengambil data dari {API_URL}. Status: {response.status_code}")
                    st.stop()

                result = response.json()

                # langsung pakai hasil transformasi dari API
                df = pd.DataFrame(result["data"])
                
                if df.empty:
                    st.warning(f"Tidak ada data dari endpoint: {API_URL}")
                    continue

                # Susun kolom akhir
                final_cols = [
                    "Tahun", "Bulan", "Overview", "Deskripsi", "Tipe", "No Akun", "Rincian Deskripsi", "Realisasi Biaya"
                ]

                df = df[final_cols].copy()

                df = df[~df["Rincian Deskripsi"].str.contains("Peny", case=False, na=False)]
                
                df_list.append(df)

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
                            
                start_dates.append(datetime.strptime(start_date, "%Y-%m-%d"))
                end_dates.append(datetime.strptime(end_date, "%Y-%m-%d"))

            st.write(df.columns.tolist())
            
            periode_awal = min(start_dates)
            periode_akhir = max(end_dates)
            
            if periode_awal.strftime("%b").lower() == periode_akhir.strftime("%b").lower():
                periode = periode_awal.strftime("%b%y").lower()
            else:
                periode = periode_awal.strftime("%b%y").lower() + "-" + periode_akhir.strftime("%b%y").lower()
            
            filename=f"prolab_{company_map.get(company_id, company_id)}_{periode}_cleaned.xlsx"

            st.caption(
                f"Perusahaan: {company_map.get(company_id, company_id).upper()} | "
                f"Periode: {periode_awal.strftime('%Y-%m-%d')} s.d. {periode_akhir.strftime('%Y-%m-%d')}"
            )
            df = pd.concat(df_list, ignore_index=True)

        st.success("Selesai!")
        st.dataframe(df.head())

        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Processing & Labour")

            workbook  = writer.book
            worksheet = writer.sheets["Processing & Labour"]

            format_angka = workbook.add_format({
                "num_format": '#,##0.00;(#,##0.00);-'
            })
            
            col_idx = df.columns.get_loc("Realisasi Biaya")
            worksheet.set_column(col_idx, col_idx, 18, format_angka)

        st.write("Company ID :", company_id)
        st.write("Start Date :", periode_awal.strftime('%Y-%m-%d'))
        st.write("End Date :", periode_akhir.strftime('%Y-%m-%d'))

        st.download_button(
            label=f"Download Hasil Cleaning: {filename}",
            data=output.getvalue(),
            file_name=filename, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_processing_labour",
            icon=":material/download:",
            type="primary"
        )