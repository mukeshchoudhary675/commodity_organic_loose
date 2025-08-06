import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Pesticide Data Processor", layout="wide")
st.title("🌿 Pesticide Parameter Processor")
st.markdown("Process Organic and Loose/Normal pesticide data from commodity files.")

def normalize_column(col):
    return re.sub(r"[^a-z]", "", str(col).lower())

def get_marker_index(df, marker):
    for idx, col in enumerate(df.columns):
        if normalize_column(col) == normalize_column(marker):
            return idx + 1  # data starts from next column
    return None

def find_column_by_name(df, search_name):
    normalized_search = normalize_column(search_name)
    for col in df.columns:
        if normalize_column(col) == normalized_search:
            return col
    return None

def extract_parameters(df, start_index):
    param_data = []
    columns = df.columns[start_index:]
    for i in range(0, len(columns), 3):
        if i + 2 >= len(columns):
            break
        param_name = columns[i]
        param_value_col = columns[i]
        compliance_col = columns[i + 1]
        limit_col = columns[i + 2]

        temp_df = df[[param_value_col, compliance_col, limit_col]].copy()
        temp_df.columns = ["Value", "Compliance", "Limit"]
        temp_df.insert(0, "Parameter", param_name)

        for col in ["Sample ID", "Sample Category", "Commodity", "Sample Type"]:
            col_match = find_column_by_name(df, col)
            if col_match:
                temp_df[col] = df[col_match]

        param_data.append(temp_df)

    return pd.concat(param_data, ignore_index=True) if param_data else pd.DataFrame()

def filter_category(df, category, category_col):
    if category_col is None:
        st.error("❌ 'Sample Category' column not found.")
        return pd.DataFrame()
    return df[df[category_col].astype(str).str.contains(category, case=False, na=False)]

def process_block(df, category, start_col, label, category_col):
    category_df = filter_category(df, category, category_col)
    if start_col is None or category_df.empty:
        st.warning(f"⚠️ Skipping block for: {label}")
        return pd.DataFrame()
    return extract_parameters(category_df, start_col)

def process_sheet(df):
    offlabel_start = get_marker_index(df, "Monitoring_off_label_pesticide_Starts")
    banned_start = get_marker_index(df, "Monitoring_banned_pesticide_Starts")
    category_col = find_column_by_name(df, "Sample Category")

    return {
        "Organic - Off-label": process_block(df, "Organic", offlabel_start, "Off-label Organic", category_col),
        "Organic - Banned": process_block(df, "Organic", banned_start, "Banned Organic", category_col),
        "Loose - Off-label": process_block(df, "Normal|Loose", offlabel_start, "Off-label Loose/Normal", category_col),
        "Loose - Banned": process_block(df, "Normal|Loose", banned_start, "Banned Loose/Normal", category_col),
    }

def to_excel_sheets(sheet_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in sheet_dict.items():
            clean_name = sheet_name[:31]  # Excel max length for sheet name
            df.to_excel(writer, sheet_name=clean_name, index=False)
    output.seek(0)
    return output

uploaded_file = st.file_uploader("📤 Upload a single commodity Excel file", type=["xlsx"])

if uploaded_file:
    sheet_names = pd.ExcelFile(uploaded_file).sheet_names
    selected_sheet = st.selectbox("🧾 Select sheet (commodity)", sheet_names)

    if selected_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        st.success(f"✅ Loaded sheet: {selected_sheet} with shape {df.shape}")

        output_dict = process_sheet(df)

        st.markdown("### 📥 Download Output")
        excel_bytes = to_excel_sheets(output_dict)
        st.download_button(
            label="📩 Download Processed Excel",
            data=excel_bytes,
            file_name=f"Processed_{selected_sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
