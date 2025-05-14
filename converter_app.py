import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Excel File Converter")

# Upload file List Order
target_file = st.file_uploader("Upload file List Order.xlsx", type=["xlsx"])

if target_file:
    # Load master dari GitHub raw
    master_df = pd.read_excel("https://raw.githubusercontent.com/keiichiro05/LO_converter/main/master.xlsx", engine='openpyxl')

    group_map = dict(zip(master_df['GROUP'], master_df['GROUP TO BE']))
    sku_map = dict(zip(master_df['SKU-LIST ORDER'], master_df['SKU TO BE']))
    lka_map = dict(zip(master_df['ship_to'], master_df['CUST_NAME']))

    target['YEAR'] = pd.to_datetime(target['po_creation_Date']).dt.year.astype(int).astype(str)

    # Load target
    target = pd.read_excel(target_file)

    # Mapping
    target['GROUP TO BE'] = target['group'].map(group_map)
    target['SKU TO BE'] = target['material_desc'].map(sku_map)
    target.loc[target['group'] == 'LKA', 'GROUP TO BE'] = target.loc[target['group'] == 'LKA', 'ship_to'].map(lka_map)

    # Konversi Bulan
    month_map = {
        'January': 'JAN', 'February': 'FEB', 'March': 'MAR',
        'April': 'APR', 'May': 'MAY', 'June': 'JUN',
        'July': 'JUL', 'August': 'AUG', 'September': 'SEP',
        'October': 'OCT', 'November': 'NOV', 'December': 'DEC'
    }
    target['MONTH'] = pd.to_datetime(target['po_creation_Date']).dt.strftime('%B').map(month_map)
    # Konversi Tahun
target['YEAR'] = pd.to_datetime(target['po_creation_Date']).dt.year.astype(int).astype(str)
target['YEAR'] = target['YEAR'].str.split('.').str[0]
str)

    # Klasifikasi SKU
    def classify_group_sku(desc):
        desc_lower = str(desc).lower()
        if 'mizone' in desc_lower:
            return 'Mizone'
        elif 'vit' in desc_lower:
            return 'VIT'
        elif desc in [
            '1500ML AQUA LOCAL MULTIPACK 1X6',
            '750ML AQUA LOCAL 1X18',
            '450ML AQUA KIDS 1X24',
            '220ML AQUA CUBE MINI BOTTLE LOCAL 1X24'
        ]:
            return 'spec SKU'
        elif desc == '1100ML AQUA LOCAL 1X12 BARCODE ON CAP':
            return 'aqua life'
        else:
            return 'sps'

    target['Group SKU'] = target['material_desc'].apply(classify_group_sku)

    # Dummy
    target['dummy'] = target['YEAR'] + ' ' + target['MONTH'] + ' ' + target['SKU TO BE'].astype(str)

    # Output DataFrame
    output = pd.DataFrame()
    output['dummy'] = target['dummy']
    output['YEAR'] = target['YEAR']
    output['MONTH'] = target['MONTH']
    output['ACCOUNT'] = target['GROUP TO BE']
    output['GROUP ACCOUNT'] = target['GROUP TO BE']
    output['SKU'] = target['SKU TO BE']
    output['material desc'] = target['material_desc']
    output['Group SKU'] = target['Group SKU']
    output['Region'] = target['region_ops']
    output['DC'] = target['dc_name_sl_forecast']
    output['PO'] = target['po_qty_cap']
    output['DO'] = target['do_qty_nett']
    output['reject_code'] = target['reject_code']
    output['sap_rejection'] = target['sap_rejection']

    # Export to Excel
    output_buffer = BytesIO()
    output.to_excel(output_buffer, index=False)
    output_buffer.seek(0)

    st.success("✅ Konversi berhasil!")
    st.download_button(
        label="⬇️ Download Hasil Excel",
        data=output_buffer,
        file_name="RAW.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
