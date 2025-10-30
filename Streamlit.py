import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import io

st.set_page_config(page_title="Estimasi Perhitungan Harga", layout="wide")
st.title("Estimasi Perhitungan Harga")

@st.cache_data
def get_data_from_google():
    with st.spinner("Getting data from Google Sheets..."):
        SCOPES = ['https://www.googleapis.com/auth/drive']

        # Authenticate using Streamlit secrets instead of local JSON
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["google_service_account"], scopes=SCOPES
        )

        client = gspread.authorize(credentials)
        sheet = client.open_by_key("16j01vWUSP8T3Dt_6ZxKJvVFSMzpjYAZp10OV887ibyA")
        worksheet = sheet.worksheet('Sheet1')

        master = pd.DataFrame(worksheet.get_all_records())
        return master

if 'master' not in st.session_state:
    st.session_state.master = get_data_from_google()
master = st.session_state.master

# --- User Inputs ---
sub_item = st.selectbox("Pilih Sub Item", master["Sub_Item"].unique())

col1, col2 = st.columns(2)
with col1:
    rmb = st.number_input("RMB", min_value=0.0, step=0.1)
    harga_kompetitor = st.number_input("Harga Kompetitor (Rp)", min_value=0, step=100)
with col2:
    konversi_beli = st.number_input("Konversi Beli", min_value=0, step=1)
    harga_retail = st.number_input("Harga Retail (Rp)", min_value=0, step=100)

# --- Look up values from df ---
selected_row = master[master["Sub_Item"] == sub_item].iloc[0]
kurs = selected_row["Kurs"]
marginlusin = selected_row["Margin_Lusin"]
marginkoli = selected_row["Margin_Koli"]
marginspecial = selected_row["Margin_Special"]
ongkir = selected_row["Ongkir"]

# --- Formula calculations ---
# 1️⃣ Harga Jual per Unit
harga_jual_per_unit_lusin = (((rmb * kurs) + (ongkir * rmb))) + ((((rmb * kurs) + (ongkir * rmb))) * marginlusin)
harga_jual_per_unit_lusin = round(harga_jual_per_unit_lusin / 500) * 500
harga_jual_per_unit_koli = (((rmb * kurs) + (ongkir * rmb))) + ((((rmb * kurs) + (ongkir * rmb))) * marginkoli)
harga_jual_per_unit_koli = round(harga_jual_per_unit_koli / 500) * 500
harga_jual_per_unit_special = (((rmb * kurs) + (ongkir * rmb))) + ((((rmb * kurs) + (ongkir * rmb))) * marginspecial)
harga_jual_per_unit_special = round(harga_jual_per_unit_special / 500) * 500

# 2️⃣ Harga Jual per Konversi
harga_jual_per_konversi_lusin = (((rmb * konversi_beli) * kurs) + (ongkir * (rmb * konversi_beli))) + \
                       ((((rmb * konversi_beli) * kurs) + (ongkir * (rmb * konversi_beli))) * marginlusin)
harga_jual_per_konversi_lusin = round(harga_jual_per_konversi_lusin / 500) * 500
harga_jual_per_konversi_koli = (((rmb * konversi_beli) * kurs) + (ongkir * (rmb * konversi_beli))) + \
                       ((((rmb * konversi_beli) * kurs) + (ongkir * (rmb * konversi_beli))) * marginkoli)
harga_jual_per_konversi_koli = round(harga_jual_per_konversi_koli / 500) * 500
harga_jual_per_konversi_special = (((rmb * konversi_beli) * kurs) + (ongkir * (rmb * konversi_beli))) + \
                       ((((rmb * konversi_beli) * kurs) + (ongkir * (rmb * konversi_beli))) * marginspecial)
harga_jual_per_konversi_special = round(harga_jual_per_konversi_special / 500) * 500

# --- Create result table ---
result_df = pd.DataFrame([{
    "Harga Kompetitor": f"Rp{harga_kompetitor:,.0f}" if harga_kompetitor else "",
    "Harga Retail": f"Rp{harga_retail:,.0f}" if harga_retail else "",
    "Sub Item": sub_item,
    "Harga Lusin per Unit": f"Rp{harga_jual_per_unit_lusin:,.0f}",
    "Harga Lusin by Konversi": f"Rp{harga_jual_per_konversi_lusin:,.0f}",
    "Harga Koli per Unit": f"Rp{harga_jual_per_unit_koli:,.0f}",
    "Harga Koli by Konversi": f"Rp{harga_jual_per_konversi_koli:,.0f}",
    "Harga Special per Unit": f"Rp{harga_jual_per_unit_special:,.0f}",
    "Harga Special by Konversi": f"Rp{harga_jual_per_konversi_special:,.0f}"
}])

st.write("### 📊 Hasil Tabel")
st.dataframe(result_df, use_container_width=True)

# Convert DataFrame to Excel in memory
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    result_df.to_excel(writer, index=False, sheet_name='Hasil')
output.seek(0)

# Download button
st.download_button(
    label="Download Excel",
    data=output,
    file_name=f"Hasil_Estimasi_{sub_item}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)



