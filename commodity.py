import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(layout="wide", page_title="Spice Pesticide Processor")

st.title("🌿 Spice Pesticide Report Generator")

uploaded_file = st.file_uploader("Upload Excel file with 13 spice sheets", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    st.success(f"✅ Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")

    selected_sheet = st.selectbox("📋 Select a Spice Sheet to Process", sheet_names)
    df = xls.parse(selected_sheet)
    df.columns = [str(col).strip() for col in df.columns]  # Normalize column names

    st.dataframe(df.head(5))

    with st.expander("🔧 Column Settings"):
        col_names = df.columns.tolist()

        commodity_col = st.selectbox("Select 'Commodity' column", col_names)
        variant_col = st.selectbox("Select 'Variant' column", col_names)

        # Auto-detect banned marker column case-insensitively
        banned_marker = next((col for col in col_names if str(col).strip().lower() == "monitoring_banned_pesticide_starts"), None)

        if banned_marker is None:
            banned_marker = st.selectbox("Select the 'Monitoring_banned_pesticide_Starts' marker column", col_names)
        else:
            st.info(f"Auto-detected banned pesticide marker column: **{banned_marker}**")

        marker_index = df.columns.get_loc(banned_marker)

        # Find first non-empty column after marker for actual banned start
        banned_start = marker_index + 1
        while banned_start < len(df.columns) and df.iloc[:, banned_start].isnull().all():
            banned_start += 1

        offlabel_start = df.columns.get_loc(df.columns[commodity_col]) + 3  # You can tweak this if needed
        offlabel_end = marker_index - 1
        banned_end = len(df.columns) - 1

    def get_parameter_indexes(start_col, end_col):
        columns = df.columns[start_col:end_col + 1]
        # Every parameter takes 3 columns: value, compliance, limit
        param_triplets = [i for i in range(start_col, end_col + 1, 3) if i + 1 <= end_col]
        return param_triplets

    def process_data(variant_filter, start_col, end_col, type_name):
        result_rows = []
        param_indexes = get_parameter_indexes(start_col, end_col)

        pesticide_data = {}
        for _, row in df.iterrows():
            commodity = row[commodity_col]
            variant = row[variant_col]

            if isinstance(variant_filter, list):
                if variant not in variant_filter:
                    continue
            else:
                if variant != variant_filter:
                    continue

            for i in param_indexes:
                pest = df.columns[i]
                value = row[i]
                compliance = str(row[i + 1]).strip().lower() if i + 1 < len(df.columns) else ""

                if pd.notna(value) and value != "":
                    try:
                        value = float(str(value).strip())
                    except (ValueError, TypeError):
                        continue

                    if pest not in pesticide_data:
                        pesticide_data[pest] = {}
                    if commodity not in pesticide_data[pest]:
                        pesticide_data[pest][commodity] = {
                            "min": None, "max": None, "total": 0, "unsafe": 0
                        }

                    rec = pesticide_data[pest][commodity]
                    rec["total"] += 1
                    if compliance == "unsafe":
                        rec["unsafe"] += 1
                        if rec["min"] is None or value < rec["min"]:
                            rec["min"] = value
                        if rec["max"] is None or value > rec["max"]:
                            rec["max"] = value

        results = [["S. No", type_name + " Pesticide Residues", "Name of Spice",
                    "Min Amount (mg/kg)", "Max Amount (mg/kg)",
                    "No. of unsafe", "Total Samples", "% Unsafe"]]

        sn = 1
        for pest, commodities in pesticide_data.items():
            for commodity, rec in commodities.items():
                percent = (rec["unsafe"] / rec["total"] * 100) if rec["total"] > 0 else 0
                results.append([
                    sn, pest, commodity,
                    rec["min"] if rec["min"] is not None else "No Residue",
                    rec["max"] if rec["max"] is not None else "No Residue",
                    rec["unsafe"], rec["total"], f"{percent:.2f}%"])
                sn += 1

        return pd.DataFrame(results[1:], columns=results[0])

    st.write("### 📊 Generate Report for Selected Commodity")

    if st.button("Generate Report"):
        organic_off = process_data("Organic", offlabel_start, offlabel_end, "Off-label Organic")
        organic_ban = process_data("Organic", banned_start, banned_end, "Banned Organic")
        normal_off = process_data(["Normal", "Loose"], offlabel_start, offlabel_end, "Off-label")
        normal_ban = process_data(["Normal", "Loose"], banned_start, banned_end, "Banned")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            organic_off.to_excel(writer, sheet_name="Organic Off-label", index=False)
            organic_ban.to_excel(writer, sheet_name="Organic Banned", index=False)
            normal_off.to_excel(writer, sheet_name="Loose Off-label", index=False)
            normal_ban.to_excel(writer, sheet_name="Loose Banned", index=False)

        st.success("✅ Report Generated Successfully!")
        st.download_button("📥 Download Final Report", output.getvalue(), "processed_spice_report.xlsx")
