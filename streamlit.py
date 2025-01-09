import pandas as pd
import requests
from io import BytesIO
from google.oauth2 import service_account
import io
import gspread
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from openpyxl.styles import NamedStyle, Alignment
import streamlit as st
import time

@st.cache_data
def get_data_from_google():
    google_creds = st.secrets["google"]
    SCOPES = ['https://www.googleapis.com/auth/drive']
    credentials = service_account.Credentials.from_service_account_info(google_creds, scopes=SCOPES)
    client = gspread.authorize(credentials)
    sheet = client.open_by_key("18t23AKiAQmK4A4dmkwqYTOGj4gNuFMEAsBpY50zJLNY")
    Foto = pd.DataFrame(sheet.worksheet('Sheet1').get_all_records())
    catalogue = pd.DataFrame(sheet.worksheet('CatalogueUpdate').get_all_records())
    return Foto, catalogue

if 'Foto' not in st.session_state:
    st.session_state.foto, st.session_state.catalogue = get_data_from_google()

Foto, catalogue = st.session_state.foto, st.session_state.catalogue
catalogue.rename(columns={'Item No.': 'ItemCode'}, inplace=True)
Foto = Foto.loc[Foto.groupby('ItemCode')['Upload Date'].idxmax()]
makepdf = pd.merge(catalogue, Foto[['ItemCode', 'Link']], on='ItemCode', how='left')
makepdf = makepdf[makepdf['Active'] == 'Y']


st.title("Catalogue Maker - Sukses Jaya")
st.divider()
st.header("How to:")
price = st.selectbox("Harga", ["Harga Under", "HargaLusin", "HargaKoli", "HargaSpecial"])
makepdf = makepdf[['Kategori', 'Sub Item', 'ItemCode', 'Link', 'Item Description', 'IsiCtn', 'Uom', 'Gudang', price]]

col1, col2, col3 = st.columns(3)
kategori = col1.multiselect("Kategori", makepdf['Kategori'].unique())
subitem = col2.multiselect("Sub Item", makepdf['Sub Item'].unique())
itemcode = col3.multiselect("Item Code", makepdf['ItemCode'].unique())
description = st.multiselect("Search by Description", makepdf['Item Description'].unique())

if kategori:
    makepdf = makepdf[makepdf['Kategori'].isin(kategori)]
if subitem:
    makepdf = makepdf[makepdf['Sub Item'].isin(subitem)]
if itemcode:
    makepdf = makepdf[makepdf['ItemCode'].isin(itemcode)]
if description:
    makepdf = makepdf[makepdf['Item Description'].isin(description)]

total_rows = len(makepdf)
st.write(f"Total Rows: {total_rows}")
st.dataframe(makepdf)

if st.button("Start"):
    with st.spinner("Membuat File Excel..."):
        start_time = time.time()
        wb = Workbook()
        ws = wb.active
        ws.title = "Data Produk"
        ws.append(makepdf.columns.tolist())
        ws.row_dimensions[1].height = 20

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20

        currency_style = NamedStyle(name="currency_style", number_format='"Rp" #,##0')

        for i, row in enumerate(dataframe_to_rows(makepdf, index=False, header=False), start=2):
            ws.append(row[:3] + [None] + row[4:])
            ws.row_dimensions[i].height = 80

            link = row[3]
            if link:
                try:
                    img = Image(BytesIO(requests.get(link).content))
                    img.width, img.height = 140, 105
                    ws.add_image(img, f"D{i}")
                except Exception:
                    pass

            ws[f"I{i}"].style = currency_style

        for col in ws.columns:
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = 20
            for cell in col:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        st.download_button(
            label="Download File Excel",
            data=buffer,
            file_name="data_produk_dengan_gambar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.write(f"File Excel berhasil dibuat dalam {time.time() - start_time:.2f} detik!")
