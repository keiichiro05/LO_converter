import pandas as pd
import streamlit as st
import io

# Fungsi untuk memuat master data
@st.cache_data
def load_master_data(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            st.error("❌ Format file master tidak didukung. Harus .xlsx atau .csv")
            return None

        # Normalisasi kolom
        df.columns = df.columns.str.strip().str.upper()
        required_columns = {'GROUP', 'GROUP TO BE', 'SKU', 'SKU TO BE'}

        customer_cols = {'CUSTOMER_NAME ', 'CUSTOMER_NAME TO BE'}
        existing_customer_cols = customer_cols & set(df.columns)

        if not required_columns.issubset(df.columns) or not existing_customer_cols:
            missing = required_columns - set(df.columns)
            if not existing_customer_cols:
                missing.add("CUSTOMER_NAME  or CUSTOMER_NAME TO BE")
            st.error(f"❌ Kolom berikut tidak ditemukan di master file: {missing}")
            return None

        return df
    except Exception as e:
        st.error(f"❌ Gagal memuat master file: {str(e)}")
        return None

# Fungsi untuk memuat list order
@st.cache_data
def load_list_order_data(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            st.error("❌ Format file List Order tidak didukung. Harus .xlsx atau .csv")
            return None

        df.columns = df.columns.str.strip().str.lower()
        required_columns = {
            'po_creation_date', 'group', 'segmen_name', 'cust_name', 'material_desc',
            'grouping_sku', 'region_ops', 'dc_name_sl_forecast',
            'po_qty_cap', 'do_qty_nett', 'reject_code', 'sap_rejection'
        }

        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            st.error(f"❌ Kolom berikut tidak ditemukan di List Order: {missing}")
            return None

        return df
    except Exception as e:
        st.error(f"❌ Gagal memuat List Order: {str(e)}")
        return None

# Fungsi untuk konversi nama bulan
def convert_month(full_month):
    if pd.isna(full_month):
        return ''
    month_map = {
        'January': 'JAN', 'February': 'FEB', 'March': 'MAR',
        'April': 'APR', 'May': 'MAY', 'June': 'JUN',
        'July': 'JUL', 'August': 'AUG', 'September': 'SEP',
        'October': 'OCT', 'November': 'NOV', 'December': 'DEC'
    }
    return month_map.get(full_month, str(full_month)[:3].upper())

# Fungsi untuk mapping GROUP ACCOUNT
def get_group_account(group_value, cust_name, customer_mapping):
    if str(group_value).strip().upper() == 'LKA':
        key = str(cust_name).strip().lower()
        return customer_mapping.get(key, 'UNKNOWN')
    else:
        return group_value

# Fungsi utama proses data
def process_data(list_order_df, master_df):
    output = pd.DataFrame()

    try:
        list_order_df['po_creation_date'] = pd.to_datetime(
            list_order_df['po_creation_date'], format='%A, %B %d, %Y', errors='coerce'
        )

        # Perbaiki material_desc spesifik
        list_order_df['material_desc'] = list_order_df['material_desc'].str.replace("''", "'", regex=False)

        output['YEAR'] = list_order_df['po_creation_date'].dt.year.fillna(0).astype(int)
        output['MONTH'] = list_order_df['po_creation_date'].dt.strftime('%B').apply(convert_month)

        # Mapping SKU
        sku_map = master_df.drop_duplicates('SKU').set_index('SKU')['SKU TO BE'].to_dict()

        # Mapping nama customer
        customer_name_key = 'CUSTOMER_NAME ' if 'CUSTOMER_NAME ' in master_df.columns else 'CUSTOMER_NAME'
        customer_name_map = {
            str(row[customer_name_key]).strip().lower(): str(row['CUSTOMER_NAME TO BE']).strip()
            for _, row in master_df.iterrows()
            if pd.notna(row.get(customer_name_key)) and pd.notna(row.get('CUSTOMER_NAME TO BE'))
        }

        output['ACCOUNT'] = list_order_df['group']
        customer_mapping = customer_name_map
        output['GROUP ACCOUNT'] = list_order_df.apply(
            lambda x: get_group_account(x['group'], x['cust_name'], customer_mapping), axis=1
        )

        # SKU to be logic with exception
        def map_sku(row):
            material = row['material_desc']
            sku_to_be = sku_map.get(material, 'UNKNOWN')

            if material == '5 GALLON AQUA LOCAL' and (row['group'].upper() == 'IGR' or row['GROUP ACCOUNT'].upper() == 'IGR'):
                return 'AQUA JUGS 19L / AQUA AIR MINERAL (BKL) GLN 19L'
            return sku_to_be

        output['SKU'] = list_order_df.assign(**output).apply(map_sku, axis=1)

        output['material desc'] = list_order_df['material_desc']
        output['Group SKU'] = list_order_df['grouping_sku']
        output['Region'] = list_order_df['region_ops']
        output['DC'] = list_order_df['dc_name_sl_forecast']
        output['PO'] = list_order_df['po_qty_cap']
        output['DO'] = list_order_df['do_qty_nett']
        output['reject_code'] = list_order_df['reject_code']
        output['sap_rejection'] = list_order_df['sap_rejection']

    except Exception as e:
        st.error(f"❌ Error saat memproses data: {str(e)}")
        return pd.DataFrame()

    return output

# ========== Streamlit UI ==========

st.set_page_config(page_title="Excel Converter", layout="wide")
st.title("📊 Excel Converter - List Order to RAW")

with st.sidebar:
    st.header("⚙️ Upload Master Data")
    master_file = st.file_uploader("Upload master.xlsx atau master.csv", type=["xlsx", "csv"])

if master_file:
    master_df = load_master_data(master_file)

    if master_df is not None:
        st.success("✅ Master file berhasil dimuat.")
        st.header("📤 Upload List Order")
        list_order_file = st.file_uploader("Upload file yang mengandung 'List Order'", type=["xlsx", "csv"])

        if list_order_file:
            if 'List Order' in list_order_file.name:
                list_order_df = load_list_order_data(list_order_file)

                if list_order_df is not None:
                    with st.spinner("🔄 Sedang memproses data..."):
                        result_df = process_data(list_order_df, master_df)

                    if not result_df.empty:
                        st.balloons()
                        st.subheader("📑 Preview Hasil Konversi")
                        st.dataframe(result_df.head())

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='RAW')

                        st.download_button(
                            label="⬇️ Download File RAW.xlsx",
                            data=output.getvalue(),
                            file_name="RAW.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("⚠️ Tidak ada data yang berhasil diproses.")
            else:
                st.error("❌ Nama file harus mengandung 'List Order'.")
else:
    st.info("📂 Silakan upload master.xlsx atau master.csv terlebih dahulu.")
