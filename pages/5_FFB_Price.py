import streamlit as st
import pandas as pd
import requests
from io import BytesIO

st.title("Get FFB Price Data")

st.info(
        "Klik **Ambil Data FFB PRICE** untuk memuat data dan mengunduh hasil dalam format Excel (.xlsx). "
    )

API_URL = "https://dashboard.mahkotagroup.com/api/dashboard/material-purchase-rp-kg/table?startMonth=2025-07"

selected_company = st.selectbox(
    "Pilih Perusahaan",
    ["BIMP", "BIMR", "BIMS", "MUL", "KPNJ"]
)

if st.button("Ambil Data FFB PRICE"):
    with st.spinner("Mengambil data dari API...."):
        response = requests.get(API_URL)

        if response.status_code != 200:
            st.error(f"Gagal mengambil data dari {API_URL}. Status: {response.status_code}")
            st.stop()

        result = response.json()

        # langsung pakai hasil transformasi dari API
        df = pd.DataFrame(result["data"])

        st.write(df.columns.tolist())
        
        # Susun kolom akhir
        final_cols = [
            "Tahun", "Bulan", "Harga TBS/Kg"
        ]
        
        df = df[df["PT"] == selected_company]

        df = df[final_cols].copy()
        
        # mengambil informasi dari respons API
        company = selected_company
        
        filename=f"FFB_Price_{company}.xlsx"

        st.caption(
            f"FFB Price | Perusahaan: {company}"
        )

    st.success("Selesai!")
    st.dataframe(df)

    # Tombol Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="FFB Price")

        workbook  = writer.book
        worksheet = writer.sheets["FFB Price"]
    
    st.write("Company :", company)
    
    st.download_button(
        label=f"Download: {filename}",
        data=output.getvalue(),
        file_name=filename, 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_ffb_price",
        icon=":material/download:",
        type="primary"
    ) 