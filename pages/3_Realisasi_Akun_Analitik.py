import streamlit as st
import pandas as pd
import numpy as np
import requests
from io import BytesIO
from datetime import datetime
from auth import require_login

require_login()

st.title("Cleaning Item Jurnal")

company_options = {
    "BIMP": [
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bimp-20250701-20251231",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bimp-20260101-20260630",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bimp-20260701-20261231"
    ],
    "BIMR": [
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bimr-20250701-20251231",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bimr-20260101-20260630",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bimr-20260701-20261231"
    ],
    "BIMS": [
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bims-20250701-20250930",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bims-20251001-20251231",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bims-20260101-20260331",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bims-20260401-20260630",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bims-20260701-20260930",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-bims-20261001-20261231"
    ],
    "MUL": [
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-mul-20250701-20251231",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-mul-20260101-20260630",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-mul-20260701-20261231"
    ],
    "KPNJ": [
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-kpnj-20250701-20251231",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-kpnj-20260101-20260331",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-kpnj-20260401-20260630",
        "https://dashboard.mahkotagroup.com/api/powerbi-feed/realisasi-kpnj-20260701-20261231"
    ]
}

source = st.radio(
    "Sumber Data",
    ["Upload File Excel", "API", "API - Custom Periode"],
    horizontal=True
)

if source == "Upload File Excel":
    st.write("*Realisasi Akun Analitik* : data **artikel jurnal** dari modul Accounting. File yang diupload harus dalam format Excel (.xlsx) dengan struktur: **:green[Item Jurnal (account.move.line) (perusahaan_periode).xlsx]**, contoh: **:green[Item Jurnal (account.move.line) (bimp_des25).xlsx]**")

    uploaded_files = st.file_uploader("Upload File Item Jurnal (Excel)", type=['xlsx'], accept_multiple_files=True)

    if not uploaded_files:
        st.warning("Silakan upload minimal 1 file")

    if uploaded_files:
        df_list = []
        periode_list = []
        
        pd.set_option("display.max_columns", None)

        with st.spinner("Sedang memproses..."):
            for uploaded_file in uploaded_files:
                # 1) Buka file Excel
                df_raw = pd.read_excel(uploaded_file)

                # 2) Ambil nama file untuk digunakan sebagai nama file hasil download
                file_name = uploaded_file.name  # contoh: Item Jurnal (account.move.line) (bimp_des25).xlsx
                file_name = file_name.replace(".xlsx", "").lower()  
                file_name = file_name.replace("item jurnal (account.move.line) (", "item_jurnal_").replace(")", "").replace(" ", "_")

                parts = file_name.split("_")  # ["item", "jurnal", "bimp", "des25"]
                prefix = "_".join(parts[:-1])

                # ambil bagian terakhir (des25)
                last = parts[-1]

                periode_list.append(last)

                # 3) Ambil hanya baris transaksi yang punya tanggal valid
                df = df_raw.copy()
                df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
                df = df[df["Tanggal"].notna()].copy()

                # 4) Pisah kolom Akun -> No Akun + Nama Akun
                akun_split = df["Akun"].astype(str).str.extract(r"^\s*(?P<No_Akun>\S+)\s*(?P<Nama_Akun>.*)$")

                df["No Akun"] = akun_split["No_Akun"].replace("nan", np.nan)
                df["Nama Akun"] = akun_split["Nama_Akun"].replace("nan", np.nan).str.strip()

                # 5) Pisah kolom Akun Analitik
                analitik = df["Akun Analitik"].astype(str).replace("nan", np.nan)

                kode_analitik = analitik.str.extract(r"^\[(?P<No_Akun_Analitik>[^\]]+)\]")
                nama_analitik = analitik.str.extract(r"^\[[^\]]+\]\s*(?P<Nama_Akun_Analitik>.*)$")

                df["No Akun Analitik"] = kode_analitik["No_Akun_Analitik"].replace("nan", np.nan).str.strip()
                df["Nama Akun Analitik"] = nama_analitik["Nama_Akun_Analitik"].replace("nan", np.nan).str.strip()

                # Kalau blank, jadikan UNDEFINED
                df["No Akun Analitik"] = df["No Akun Analitik"].fillna("UNDEFINED")
                df["Nama Akun Analitik"] = df["Nama Akun Analitik"].fillna("UNDEFINED")

                df["No Akun Analitik"] = df["No Akun Analitik"].replace(r"^\s*$", "UNDEFINED", regex=True)
                df["Nama Akun Analitik"] = df["Nama Akun Analitik"].replace(r"^\s*$", "UNDEFINED", regex=True)

                # Rapikan dulu kalau ada yang nempel seperti 008-10(B).003.015
                df["No Akun Analitik"] = df["No Akun Analitik"].str.replace(r"(?<=\d)\(", " (", regex=True)

                # Kode Induk Analitik
                # Contoh:
                # 008-10.027.002.16 -> 008-10
                # 008-09.001.001 -> 008-09
                # 008-09 (A).001.001 -> 008-09 (A)
                df["Kode Induk Analitik"] = (
                    df["No Akun Analitik"]
                    .str.extract(r"^([^.]+(?:\s*\([^)]+\))?)\.", expand=False)
                    .str.strip()
                )

                # Kode Detail Analitik
                # Contoh:
                # 008-10.027.002.16 -> 027.002.16
                # 008-09.001.001 -> 001.001
                # 008-09 (A).001.001 -> 001.001
                df["Kode Detail Analitik"] = (
                    df["No Akun Analitik"]
                    .str.extract(r"^[^.]+(?:\s*\([^)]+\))?\.(.+)$", expand=False)
                    .str.strip()
                )

                # Tipe Unit -> ambil dari bagian sebelum ":" pertama
                df["Tipe Unit"] = (
                    df["Nama Akun Analitik"]
                    .astype(str)
                    .str.extract(r"^(.*?)\s*:", expand=False)
                    .str.extract(r"(?:\(([^)]+)\)|([A-Z]+))\s*$")
                    .bfill(axis=1)
                    .iloc[:, 0]
                    .str.strip()
                )

                # Kalau hasil turunan kosong, jadikan UNDEFINED
                df["Kode Induk Analitik"] = df["Kode Induk Analitik"].fillna("UNDEFINED")
                df["Kode Detail Analitik"] = df["Kode Detail Analitik"].fillna("UNDEFINED")
                df["Tipe Unit"] = df["Tipe Unit"].fillna("UNDEFINED")

                # 6) Buat kolom Nominal
                # Debit tetap positif
                # Kredit menjadi negatif
                df["Debit"] = pd.to_numeric(df["Debit"], errors="coerce").fillna(0)
                df["Kredit"] = pd.to_numeric(df["Kredit"], errors="coerce").fillna(0)

                df["Nominal"] = df["Debit"] - df["Kredit"]

                # 7) Filter No Akun yang tidak diinginkan (01, 04, 11, dll sesuai notebook)
                remove_prefixes = ("01", "04", "11", "12", "21", "22", "31", "32", "34", "91", "92", "99")
                df = df[~df["No Akun"].fillna("").str.startswith(remove_prefixes)].copy()

                # 8) Tambah Tahun dan Bulan
                bulan_map = {
                    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
                    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
                    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
                }

                df["Tahun"] = df["Tanggal"].dt.year
                df["Bulan"] = df["Tanggal"].dt.month.map(bulan_map)
                df["Tanggal"] = df["Tanggal"].dt.strftime("%d/%m/%Y")

                # 9) Tambah Kolom Deskripsi
                df["Deskripsi"] = pd.Series(index=df.index, dtype="object")

                text_source = (
                    df["Nama Akun"].fillna("") + " " +
                    df.get("Label", pd.Series(index=df.index, dtype="object")).fillna("")
                )

                df.loc[text_source.str.contains("Pembelian", case=False, na=False), "Deskripsi"] = "Pembelian"
                df.loc[text_source.str.contains("Harga Pokok", case=False, na=False), "Deskripsi"] = "Persediaan Awal"
                df.loc[text_source.str.contains("Transit", case=False, na=False), "Deskripsi"] = "Persediaan Akhir"
                df.loc[text_source.str.contains("Transfer", case=False, na=False), "Deskripsi"] = "Transfer Internal"

                df.loc[df["No Akun"].fillna("").str.startswith("41"), "Deskripsi"] = "Pendapatan"
                df.loc[df["No Akun"].fillna("").str.startswith("51.02"), "Deskripsi"] = "Biaya Pembelian"
                df.loc[df["No Akun"].fillna("").str.startswith("51.03"), "Deskripsi"] = "Biaya Gaji"
                df.loc[df["No Akun"].fillna("").str.startswith("52"), "Deskripsi"] = "Biaya Overhead"
                df.loc[df["No Akun"].fillna("").str.startswith("62"), "Deskripsi"] = "Beban Penjualan"
                df.loc[df["No Akun"].fillna("").str.startswith("70"), "Deskripsi"] = "Beban Administrasi"
                df.loc[df["No Akun"].fillna("").str.startswith("80"), "Deskripsi"] = "Pendapatan Lain-lain"
                df.loc[df["No Akun"].fillna("").str.startswith("81"), "Deskripsi"] = "Beban Lain-lain"

                # Override paling akhir
                df.loc[df["No Akun"].eq("41.01.01.0000.11"), "Deskripsi"] = "Pendapatan Lain-lain"
                df.loc[df["No Akun"].eq("81.07.00.0000.02"), "Deskripsi"] = "Pendapatan Lain-lain"

                # Balik tanda nominal untuk Pendapatan dan Pendapatan Lain-lain
                mask_pendapatan = df["Deskripsi"].isin(["Pendapatan", "Pendapatan Lain-lain"])
                df.loc[mask_pendapatan, "Nominal"] = df.loc[mask_pendapatan, "Nominal"] * -1

                # 10) Buat kolom Overview dari Deskripsi
                mapping = {
                    "Pendapatan/Penjualan": ["Pendapatan"],
                    "Beban Pokok Pendapatan/COGS": [
                        "Pembelian", "Persediaan Awal", "Persediaan Akhir",
                        "Transfer Internal", "Biaya Pembelian", "Biaya Gaji", "Biaya Overhead"
                    ],
                    "Beban Operasional": ["Beban Penjualan", "Beban Administrasi"],
                    "Pendapatan Non-Operasional": ["Pendapatan Lain-lain"],
                    "Beban Non-Operasional": ["Beban Lain-lain"]
                }

                overview_map = {desc: ov for ov, desc_list in mapping.items() for desc in desc_list}
                df["Overview"] = df["Deskripsi"].map(overview_map)

                df[["Deskripsi", "Overview"]].drop_duplicates().sort_values(["Overview", "Deskripsi"])

                # 11) Susun kolom akhir
                final_cols = [
                    "Tanggal", "Tahun", "Bulan", "Overview", "Deskripsi",
                    "No Akun", "Nama Akun",
                    "No Akun Analitik", "Nama Akun Analitik",
                    "Kode Induk Analitik", "Kode Detail Analitik", "Tipe Unit",
                    "Nominal"
                ]

                df_final = df[final_cols].copy()

                df_list.append(df_final)

            # gabungkan semua
            df_final_all = pd.concat(df_list, ignore_index=True)

            st.success("Selesai!")
            st.dataframe(df_final_all.head())

            # Tombol Download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final_all.to_excel(writer, index=False, sheet_name="Item Jurnal Cleaned")

                workbook  = writer.book
                worksheet = writer.sheets["Item Jurnal Cleaned"]

                format_angka = workbook.add_format({
                    "num_format": '#,##0.00;(#,##0.00);-'
                })

                col_idx = df_final_all.columns.get_loc("Nominal")
                worksheet.set_column(col_idx, col_idx, 18, format_angka)

            # Kondisi untuk filename berdasarkan jumlah periode yang diproses
            if len(periode_list) == 1:
                periode_str = periode_list[0]
            else:            
                periode_str = f"{periode_list[0]}-{periode_list[-1]}"
            
            st.download_button(
                label=f"Download {prefix}_{periode_str}_cleaned.xlsx",
                data=output.getvalue(),
                file_name=f"{prefix}_{periode_str}_cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_cleaned_item_jurnal",
                icon=":material/download:",
                type="primary"
            )

elif source == "API":
    st.info(
        "Data artikel jurnal diambil langsung dari API Dashboard. "
        "Klik **Ambil Data Artikel Jurnal** untuk memuat data dan mengunduh hasil dalam format Excel (.xlsx). "
        "Nama file akan dibuat otomatis berdasarkan perusahaan dan periode, misalnya "
        "`item_jurnal_bimp_jan26-mei26_cleaned.xlsx`."
    )

    selected_company = st.selectbox(
        "Pilih Perusahaan",
        list(company_options.keys())
    )

    API_URLS = company_options[selected_company]

    if st.button("Ambil Data Artikel Jurnal"):
        with st.spinner("Mengambil data dari API...."):
            df_list = []
            start_dates = []
            end_dates = []
            
            for API_URL in API_URLS:
                response = requests.get(API_URL)

                if response.status_code != 200:
                    st.error(f"Gagal mengambil data. Status: {response.status_code}")
                    st.stop()

                result = response.json()

                # langsung pakai hasil transformasi dari API
                df = pd.DataFrame(result["data"])
                
                if df.empty:
                    st.warning(f"Tidak ada data dari endpoint: {API_URL}")
                    continue

                # Balik tanda nominal untuk Pendapatan dan Pendapatan Lain-lain
                mask_pendapatan = df["Deskripsi"].isin(["Pendapatan", "Pendapatan Lain-lain"])
                df.loc[mask_pendapatan, "Nominal"] = df.loc[mask_pendapatan, "Nominal"] * -1

                # format tanggal
                df["Tanggal"] = pd.to_datetime(df["Tanggal"])
                df["Tanggal"] = df["Tanggal"].dt.strftime("%d/%m/%Y")

                # Susun kolom akhir
                final_cols = [
                    "Tanggal", "Tahun", "Bulan", "Overview", "Deskripsi",
                    "No Akun", "Nama Akun",
                    "No Akun Analitik", "Nama Akun Analitik",
                    "Kode Induk Analitik", "Kode Detail Analitik", "Tipe Unit",
                    "Nominal"
                ]

                df = df[final_cols].copy()
                
                # df digabung ke dalam list
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
            
            periode = periode_awal.strftime("%d %b %y").lower() + " - " + periode_akhir.strftime("%d %b %y").lower()
            
            filename=f"item_jurnal_terekam_{company_map.get(company_id, company_id)}_{periode}_cleaned.xlsx"

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
            df.to_excel(writer, index=False, sheet_name="Item Jurnal Cleaned")

            workbook  = writer.book
            worksheet = writer.sheets["Item Jurnal Cleaned"]

            format_angka = workbook.add_format({
                "num_format": '#,##0.00;(#,##0.00);-'
            })
            
            col_idx = df.columns.get_loc("Nominal")
            worksheet.set_column(col_idx, col_idx, 18, format_angka)

        st.write("Company ID :", company_id)
        st.write("Start Date :", periode_awal.strftime('%Y-%m-%d'))
        st.write("End Date :", periode_akhir.strftime('%Y-%m-%d'))

        st.download_button(
            label=f"Download Hasil Cleaning: {filename}",
            data=output.getvalue(),
            file_name=filename, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cleaned_item_jurnal_terekam",
            icon=":material/download:",
            type="primary"
        )
        
elif source == "API - Custom Periode":
    st.info(
        "Data artikel jurnal diambil dari endpoint API yang sama, "
        "kemudian difilter berdasarkan tanggal yang dipilih."
    )

    selected_company = st.selectbox(
        "Pilih Perusahaan",
        list(company_options.keys())
    )

    col1, col2 = st.columns(2)

    with col1:
        tanggal_mulai = st.date_input(
            "Tanggal Mulai"
        )

    with col2:
        tanggal_akhir = st.date_input(
            "Tanggal Akhir"
        )

    if tanggal_mulai > tanggal_akhir:
        st.error(
            "Tanggal mulai tidak boleh lebih besar "
            "dari tanggal akhir."
        )
        st.stop()

    if st.button("Ambil Data Artikel Jurnal"):

        API_URLS = company_options[selected_company]

        with st.spinner("Mengambil data dari API..."):

            df_list = []

            for API_URL in API_URLS:

                response = requests.get(API_URL)

                if response.status_code != 200:
                    st.error(
                        f"Gagal mengambil data. "
                        f"Status: {response.status_code}"
                    )
                    st.stop()

                result = response.json()

                df = pd.DataFrame(result["data"])

                if df.empty:
                    continue
                
                # Ubah tanggal
                df["Tanggal"] = pd.to_datetime(
                    df["Tanggal"],
                    errors="coerce"
                )

                # filter sesuai input user
                df = df[
                    (df["Tanggal"].dt.date >= tanggal_mulai) &
                    (df["Tanggal"].dt.date <= tanggal_akhir)
                ].copy()
                
                if df.empty:
                    continue

                # Lanjutkan cleaning yang sudah ada
                mask_pendapatan = df["Deskripsi"].isin([
                    "Pendapatan",
                    "Pendapatan Lain-lain"
                ])

                df.loc[mask_pendapatan, "Nominal"] *= -1

                df["Tanggal"] = df["Tanggal"].dt.strftime(
                    "%d/%m/%Y"
                )

                final_cols = [
                    "Tanggal", "Tahun", "Bulan",
                    "Overview", "Deskripsi",
                    "No Akun", "Nama Akun",
                    "No Akun Analitik",
                    "Nama Akun Analitik",
                    "Kode Induk Analitik",
                    "Kode Detail Analitik",
                    "Tipe Unit",
                    "Nominal"
                ]

                df_list.append(df[final_cols].copy())

            if not df_list:
                st.warning(
                    "Tidak ada data pada tanggal yang dipilih."
                )
                st.stop()

            df = pd.concat(
                df_list,
                ignore_index=True
            )

        st.success("Selesai!")

        st.dataframe(df.head())
        
        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Item Jurnal Cleaned")

            workbook  = writer.book
            worksheet = writer.sheets["Item Jurnal Cleaned"]

            format_angka = workbook.add_format({
                "num_format": '#,##0.00;(#,##0.00);-'
            })
            
            col_idx = df.columns.get_loc("Nominal")
            worksheet.set_column(col_idx, col_idx, 18, format_angka)

        st.write("Company :", selected_company)
        st.write("Start Date :", tanggal_mulai.strftime('%Y-%m-%d'))
        st.write("End Date :", tanggal_akhir.strftime('%Y-%m-%d'))
        
        periode =tanggal_mulai.strftime("%d %b %y").lower() + " - " + tanggal_akhir.strftime("%d %b %y").lower()
                    
        filename=f"item_jurnal_terekam_{selected_company.lower()}_{periode}_cleaned.xlsx"

        st.download_button(
            label=f"Download Hasil Cleaning: {filename}",
            data=output.getvalue(),
            file_name=filename, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cleaned_item_jurnal_terekam_custom",
            icon=":material/download:",
            type="primary"
        )