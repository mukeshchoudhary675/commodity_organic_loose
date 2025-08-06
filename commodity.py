import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("Spice Pesticide Processor - Organic & Loose")

uploaded_file = st.file_uploader("Upload Excel file with 1 commodity sheet", type=[".xlsx"])

if uploaded_file:
    # Load Excel file
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = xls.parse(sheet)

    st.success(f"Loaded sheet: {sheet} with {df.shape[0]} rows and {df.shape[1]} columns")

    # Normalize columns (case insensitive)
    df.columns = [str(c).strip() for c in df.columns]

    # Dropdowns for column selection
    commodity_col = st.selectbox("Select column for 'Commodity'", df.columns)
    variant_col = st.selectbox("Select column for 'Variant'", df.columns)

    # Find banned marker column (case-insensitive match)
    banned_marker_col = next((col for col in df.columns if re.match(r"monitoring_banned_pesticide_starts", str(col), re.IGNORECASE)), None)

    if not banned_marker_col:
        st.error("❌ Could not find 'Monitoring_banned_pesticide_Starts' column.")
    else:
        banned_start_index = df.columns.get_loc(banned_marker_col)

        # If banned_marker_col is empty, skip to next column
        if df[banned_marker_col].isnull().all() and banned_start_index + 1 < len(df.columns):
            banned_start_index += 1

        # Determine number of parameters
        total_params = (len(df.columns) - (banned_start_index)) // 3
        banned_end_index = banned_start_index + total_params * 3

        offlabel_start_index = banned_start_index - total_params * 3 if banned_start_index - total_params * 3 >= 0 else 0
        offlabel_end_index = banned_start_index

        def process_subset(df, variant_values, start, end):
            filtered = df[df[variant_col].str.strip().str.lower().isin([v.lower() for v in variant_values])]
            core = filtered[[commodity_col, variant_col]].copy()

            param_df = filtered.iloc[:, start:end]

            # Only include columns that are part of parameter sets (every 3 cols)
            valid_param_cols = param_df.columns[: (param_df.shape[1] // 3) * 3]
            return pd.concat([core, param_df[valid_param_cols]], axis=1)

        # ORGANIC
        organic_off = process_subset(df, ["Organic"], offlabel_start_index, offlabel_end_index)
        organic_ban = process_subset(df, ["Organic"], banned_start_index, banned_end_index)

        # LOOSE + NORMAL
        loose_off = process_subset(df, ["Loose", "Normal"], offlabel_start_index, offlabel_end_index)
        loose_ban = process_subset(df, ["Loose", "Normal"], banned_start_index, banned_end_index)

        # Prepare Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            organic_off.to_excel(writer, index=False, sheet_name="Organic Off-label")
            organic_ban.to_excel(writer, index=False, sheet_name="Organic Banned")
            loose_off.to_excel(writer, index=False, sheet_name="Loose Off-label")
            loose_ban.to_excel(writer, index=False, sheet_name="Loose Banned")

        st.download_button(
            label="📥 Download Processed Excel",
            data=output.getvalue(),
            file_name=f"Processed_{sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
