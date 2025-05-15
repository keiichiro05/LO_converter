import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Excel/CSV File Converter for List Order to RAW")

# Fungsi untuk memuat master data awal
@st.cache_data
def load_initial_master():
    try:
        return pd.read_excel("https://raw.githubusercontent.com/keiichiro05/LO_converter/main/master.xlsx", engine='openpyxl')
    except:
        # Jika gagal, buat DataFrame kosong dengan kolom yang diperlukan
        return pd.DataFrame(columns=[
            'GROUP', 'GROUP TO BE', 'SKU-LIST ORDER', 
            'SKU TO BE', 'ship_to', 'CUST_NAME'
        ])

# Inisialisasi master data di session state jika belum ada
if 'master_data' not in st.session_state:
    st.session_state.master_data = load_initial_master()

# Section untuk edit master data
with st.expander("✏️ Edit Master Data", expanded=False):
    st.write("Edit master data langsung di tabel berikut:")
    
    # Editor data
    edited_master = st.data_editor(
        st.session_state.master_data,
        num_rows="dynamic",  # Memungkinkan menambah/hapus baris
        use_container_width=True,
        height=400
    )
    
    # Tombol untuk menyimpan perubahan
    if st.button("💾 Save Master Data Changes"):
        st.session_state.master_data = edited_master
        st.success("Master data updated successfully!")
        
    # Tombol untuk menambah kolom jika diperlukan
    if st.button("➕ Add New Column"):
        new_col_name = st.text_input("New column name")
        if new_col_name:
            st.session_state.master_data[new_col_name] = ""
            st.rerun()

# Gunakan master data dari session state
master_df = st.session_state.master_data

if master_df.empty:
    st.warning("Master data is empty. Please add data to continue.")
    st.stop()

# Lanjutkan dengan pemetaan data seperti sebelumnya
group_map = dict(zip(master_df['GROUP'], master_df['GROUP TO BE']))
sku_map = dict(zip(master_df['SKU-LIST ORDER'], master_df['SKU TO BE']))
lka_map = dict(zip(master_df['ship_to'], master_df['CUST_NAME']))


        # Tentukan apakah file yang diupload Excel atau CSV
        if uploaded_file.name.endswith('.xlsx'):
            target = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            target = pd.read_csv(uploaded_file)
        
        # Pastikan kolom yang dibutuhkan ada di data
        if 'po_creation_Date' not in target.columns or 'group' not in target.columns or 'material_desc' not in target.columns:
            st.error("❌ File tidak memiliki kolom yang dibutuhkan!")
        else:
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
            target['YEAR'] = pd.to_datetime(target['po_creation_Date']).dt.year.astype(str)

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

            # Export ke Excel atau CSV
            file_extension = uploaded_file.name.split('.')[-1]
            output_buffer = BytesIO()

            if file_extension == "xlsx":
                output.to_excel(output_buffer, index=False)
            elif file_extension == "csv":
                output.to_csv(output_buffer, index=False)

            output_buffer.seek(0)

            st.success("✅ Konversi berhasil!")
            st.download_button(
                label="⬇️ Download Hasil",
                data=output_buffer,
                file_name=f"RAW.{file_extension}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_extension == "xlsx" else "text/csv"
            )
    else:
        st.error("❌ Nama file harus mengandung 'List Order'.")
