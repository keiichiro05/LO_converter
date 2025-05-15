import pandas as pd
import streamlit as st
from io import BytesIO

# Function to load initial master data
@st.cache_data
def load_initial_master():
    try:
        return pd.read_excel("https://raw.githubusercontent.com/keiichiro05/LO_converter/main/master.xlsx", engine='openpyxl')
    except:
        # Fallback empty DataFrame with required columns
        return pd.DataFrame(columns=[
            'GROUP', 'GROUP TO BE', 'SKU-LIST ORDER', 
            'SKU TO BE', 'ship_to', 'CUST_NAME'
        ])

# Initialize session state for master data
if 'master_data' not in st.session_state:
    st.session_state.master_data = load_initial_master()

# App title
st.title("Excel/CSV File Converter for List Order to RAW")

# Master Data Editor Section
with st.expander("✏️ Edit Master Data", expanded=False):
    st.write("Edit the master mapping table directly below:")
    
    # Data editor widget
    edited_master = st.data_editor(
        st.session_state.master_data,
        num_rows="dynamic",
        use_container_width=True,
        height=400,
        key="master_editor"
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("💾 Save Changes", help="Save changes to master data for this session"):
            st.session_state.master_data = edited_master
            st.success("Master data updated successfully!")
            
    with col2:
        if st.button("🔄 Reset Changes", help="Revert to last saved version"):
            st.rerun()
    
    # Add new column feature
    new_col_name = st.text_input("Add new column:")
    if st.button("➕ Add Column") and new_col_name:
        if new_col_name not in st.session_state.master_data.columns:
            st.session_state.master_data[new_col_name] = ""
            st.rerun()
        else:
            st.warning(f"Column '{new_col_name}' already exists!")

# Get the master data from session state
master_df = st.session_state.master_data

# Check if master data has required columns
required_columns = ['GROUP', 'GROUP TO BE', 'SKU-LIST ORDER', 'SKU TO BE', 'ship_to', 'CUST_NAME']
if not all(col in master_df.columns for col in required_columns):
    st.error("Master data is missing required columns! Please edit the master data.")
    st.stop()

# Create mapping dictionaries
group_map = dict(zip(master_df['GROUP'], master_df['GROUP TO BE']))
sku_map = dict(zip(master_df['SKU-LIST ORDER'], master_df['SKU TO BE']))
lka_map = dict(zip(master_df['ship_to'], master_df['CUST_NAME']))

# File Upload and Conversion Section
uploaded_file = st.file_uploader("Upload file List Order", type=["xlsx", "csv"])

if uploaded_file:
    # Check if filename contains 'List Order'
    if "List Order" not in uploaded_file.name:
        st.error("❌ Nama file harus mengandung 'List Order'.")
        st.stop()
    
    try:
        # Read the uploaded file
        if uploaded_file.name.endswith('.xlsx'):
            target = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            target = pd.read_csv(uploaded_file)
        
        # Validate required columns in the uploaded file
        required_upload_columns = ['po_creation_Date', 'group', 'material_desc', 'ship_to', 'region_ops', 
                                  'dc_name_sl_forecast', 'po_qty_cap', 'do_qty_nett', 'reject_code', 'sap_rejection']
        
        missing_columns = [col for col in required_upload_columns if col not in target.columns]
        if missing_columns:
            st.error(f"❌ File tidak memiliki kolom yang dibutuhkan: {', '.join(missing_columns)}")
            st.stop()
        
        # Apply mappings
        target['GROUP TO BE'] = target['group'].map(group_map)
        target['SKU TO BE'] = target['material_desc'].map(sku_map)
        target.loc[target['group'] == 'LKA', 'GROUP TO BE'] = target.loc[target['group'] == 'LKA', 'ship_to'].map(lka_map)

        # Convert dates
        month_map = {
            'January': 'JAN', 'February': 'FEB', 'March': 'MAR',
            'April': 'APR', 'May': 'MAY', 'June': 'JUN',
            'July': 'JUL', 'August': 'AUG', 'September': 'SEP',
            'October': 'OCT', 'November': 'NOV', 'December': 'DEC'
        }
        target['MONTH'] = pd.to_datetime(target['po_creation_Date']).dt.strftime('%B').map(month_map)
        target['YEAR'] = pd.to_datetime(target['po_creation_Date']).dt.year.astype(str)

        # Classify SKU groups
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

        # Create output DataFrame
        output = pd.DataFrame()
        output['dummy'] = target['YEAR'] + ' ' + target['MONTH'] + ' ' + target['SKU TO BE'].astype(str)
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

        # Prepare file for download
        file_extension = uploaded_file.name.split('.')[-1]
        output_buffer = BytesIO()

        if file_extension == "xlsx":
            output.to_excel(output_buffer, index=False)
        elif file_extension == "csv":
            output.to_csv(output_buffer, index=False)

        output_buffer.seek(0)

        # Show success and download button
        st.success("✅ Konversi berhasil!")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️ Download Hasil",
                data=output_buffer,
                file_name=f"RAW.{file_extension}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_extension == "xlsx" else "text/csv"
            )
        
        with col2:
            if st.button("👀 Preview Output"):
                st.dataframe(output.head(), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
