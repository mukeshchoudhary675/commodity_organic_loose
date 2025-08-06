import pandas as pd

def find_column_index(columns, marker):
    marker = marker.strip().lower()
    for i, col in enumerate(columns):
        if col and col.strip().lower() == marker:
            return i
    return -1

def extract_parameter_blocks(df, start_col_idx):
    # Extract all columns after start_col_idx in blocks of 3 (value, compliance, limit)
    columns = df.columns[start_col_idx + 1:]
    total_cols = len(columns)
    param_blocks = []

    for i in range(0, total_cols, 3):
        block = columns[i:i+3]
        if len(block) == 3:
            param_blocks.append(block)

    return param_blocks

def filter_by_category(df, category):
    return df[df["Category"].str.lower().isin([c.lower() for c in category])]

def build_dataset(df, param_blocks, category):
    filtered_df = filter_by_category(df, category)
    selected_cols = []

    for block in param_blocks:
        selected_cols.extend(block)

    meta_cols = [col for col in df.columns if col not in selected_cols]
    result = filtered_df[meta_cols + selected_cols]
    return result

def process_sheet(df, banned_marker, offlabel_marker):
    # Normalize columns
    df.columns = [str(col).strip() for col in df.columns]

    # --- Banned Pesticides ---
    banned_start = find_column_index(df.columns, banned_marker)
    banned_params = extract_parameter_blocks(df, banned_start)

    # --- Off-label Pesticides ---
    offlabel_start = find_column_index(df.columns, offlabel_marker)
    offlabel_params = extract_parameter_blocks(df, offlabel_start)

    # Process datasets
    data = {
        "Organic - Off-label": build_dataset(df, offlabel_params, ["Organic"]),
        "Organic - Banned": build_dataset(df, banned_params, ["Organic"]),
        "Loose - Off-label": build_dataset(df, offlabel_params, ["Loose", "Normal"]),
        "Loose - Banned": build_dataset(df, banned_params, ["Loose", "Normal"])
    }

    return data

def save_output(data, commodity_name, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for sheet_name, df in data.items():
            safe_sheet_name = f"{sheet_name} {commodity_name}"[:31]
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
