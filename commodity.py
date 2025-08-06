import streamlit as st
import pandas as pd
import re
from io import BytesIO

def get_column_index_by_name(columns, target):
    """Find the index of a column name, case-insensitive and trimmed."""
    for i, col in enumerate(columns):
        if col and isinstance(col, str) and col.strip().lower() == target.strip().lower():
            return i
    return None

def extract_parameter_blocks(columns, start_index, end_index=None):
    """Extract parameter blocks of 3 columns each: value, compliance, limit."""
    param_blocks = []
    step = 3
    end_index = end_index or len(columns)
    for i in range(start_index, end_index, step):
        if i + 2 < len(columns):
            param_name = columns[i]
            if pd.isna(param_name) or not isinstance(param_name, str) or param_name.strip() == "":
                continue
            param_blocks.append((i, columns[i:i+3]))
    return param_blocks

def process_category(df, category_value, variant_value, banned_marker):
    df_filtered = df[
        df['Category'].astype(str).str.strip().str.lower() == category_value.strip().lower()
    ]
    if variant_value:
        df_filtered = df_filtered[
            df_filtered['Variant'].astype(str).str.strip().str.lower() == variant_value.strip().lower()
        ]

    df_filtered = df_filtered.reset_index(drop=True)
    headers = list(df_filtered.columns)
    banned_marker_index = get_column_index_by_name(headers, banned_marker)

    if banned_marker_index is None:
        raise ValueError(f"Could not find banned pesticide marker column: {banned_marker}")

    # Off-label: all before banned_marker
    off_label_blocks = extract_parameter_blocks(headers, start_index=headers.index(headers[0]), end_index=banned_marker_index + 1)
    # Banned: all after banned_marker
    banned_blocks = extract_parameter_blocks(headers, start_index=banned_marker_index + 1)

    def build_df(blocks):
        data = []
        for idx, row in df_filtered.iterrows():
            for col_idx, names in blocks:
                param_name = names[0]
                value = row.iloc[col_idx]
                compliance = row.iloc[col_idx + 1]
                limit = row.iloc[col_idx + 2]
                data.append({
                    "Sample ID": row.get("Sample ID", idx + 1),
                    "Pesticide": param_name,
                    "Value": value,
                    "Compliance": compliance,
                    "Limit": limit
                })
        return pd.DataFrame(data)

    return build_df(off_label_blocks), build_df(banned_blocks)

def generate_excel(organic_off, organic_ban, loose_off, loose_ban):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        organic_off.to_excel(writer, sheet_name="Organic_Off_Label", index=False)
        organic_ban.to_excel(writer, sheet_name="Organic_Banned", index=False)
        loose_off.to_excel(writer, sheet_name="Loose_Off_Label", index=False)
        loose_ban.to_excel(writer, sheet_name="Loose_Banned", index=False)
    output.seek(0)
    return output

# === Streamlit UI ===
st.title("🧪 Pesticide Parameter Extractor")

uploaded_file = st.file_uploader("Upload Excel File with Sheets for Each Commodity", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("Select Commodity Sheet", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet)

    category_col = st.selectbox("Select Category Column", df.columns)
    variant_col = st.selectbox("Select Variant Column", df.columns)
    unique_categories = df[category_col].dropna().unique()
    category = st.selectbox("Select Category", unique_categories)

    unique_variants = df[variant_col].dropna().unique()
    variant = st.selectbox("Select Variant (Optional)", [""] + list(unique_variants))

    marker_input = st.text_input("Enter Banned Pesticide Marker Column", value="Monitoring_banned_pesticide_Starts")

    if st.button("Process and Download Excel"):
        df = df.rename(columns=lambda x: str(x).strip())  # normalize headers
        df = df.rename(columns={category_col: "Category", variant_col: "Variant"})  # map to expected keys

        try:
            organic_off, organic_ban = process_category(df, category, variant, marker_input)
            loose_off, loose_ban = process_category(df, category, variant, marker_input)
            excel_bytes = generate_excel(organic_off, organic_ban, loose_off, loose_ban)

            st.success("✅ Excel generated successfully!")
            st.download_button("📥 Download Excel File", data=excel_bytes, file_name=f"{category}_{variant}_Pesticide_Report.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")
