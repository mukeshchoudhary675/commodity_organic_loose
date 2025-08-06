import streamlit as st
import pandas as pd
import re
from io import BytesIO

def normalize_column(col):
    return re.sub(r"[^a-z]", "", str(col).lower())

def get_marker_index(df, marker):
    for i, col in enumerate(df.columns):
        if normalize_column(col) == normalize_column(marker):
            return i + 1  # Start from next column
    return None

def extract_parameters(df, start_index):
    param_cols = df.columns[start_index:]
    num_params = len(param_cols) // 3
    param_data = []

    for i in range(num_params):
        value_col = param_cols[i*3]
        compliance_col = param_cols[i*3 + 1]
        limit_col = param_cols[i*3 + 2]

        param_name = re.sub(r"_value$|_compliance$|_limit$", "", str(value_col), flags=re.IGNORECASE)

        subset = df[["Sample ID", value_col, compliance_col, limit_col]].copy()
        subset.columns = ["Sample ID", "Value", "Compliance", "Limit"]
        subset["Parameter"] = param_name
        param_data.append(subset)

    return pd.concat(param_data, ignore_index=True)

def filter_category(df, category):
    return df[df["Sample Category"].str.contains(category, case=False, na=False)]

def process_block(df, category, start_col, label):
    category_df = filter_category(df, category)
    if start_col is None:
        st.warning(f"Start marker not found for {label}. Skipping...")
        return pd.DataFrame()
    return extract_parameters(category_df, start_col)

def process_sheet(df):
    offlabel_start = get_marker_index(df, "Monitoring_off_label_pesticide_Starts")
    banned_start = get_marker_index(df, "Monitoring_banned_pesticide_Starts")

    return {
        "Organic - Off-label": process_block(df, "Organic", offlabel_start, "Off-label Organic"),
        "Organic - Banned": process_block(df, "Organic", banned_start, "Banned Organic"),
        "Loose - Off-label": process_block(df, "Normal|Loose", offlabel_start, "Off-label Loose/Normal"),
        "Loose - Banned": process_block(df, "Normal|Loose", banned_start, "Banned Loose/Normal"),
    }

def to_excel(output_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in output_dict.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    output.seek(0)
    return output

# Streamlit UI
st.title("🧪 Pesticide Parameter Processor")

uploaded_file = st.file_uploader("Upload Excel File with All Commodities", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    st.success("Excel file uploaded.")

    sheet_name = st.selectbox("Select a commodity (sheet):", xls.sheet_names)

    if st.button("Process Selected Sheet"):
        df = xls.parse(sheet_name)
        output_dict = process_sheet(df)

        st.success("✅ Processing complete!")

        output_excel = to_excel(output_dict)

        st.download_button(
            label="📥 Download Processed Excel",
            data=output_excel,
            file_name=f"Processed_{sheet_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
