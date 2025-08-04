import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Constants
TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH = "Mapping - Automation.xlsx"

# Load mapping sheet
@st.cache_data
def load_mapping():
    df = pd.read_excel(MAPPING_PATH, sheet_name=0)
    return df

def process_mapping_mode(input_file, mapping_df):
    input_df = pd.read_excel(input_file)
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_values = wb["Values"]
    ws_type = wb["Types"]

    # Create a new dict for output columns
    col_data_map = {}

    # Step 1: Keep original columns
    for col in input_df.columns:
        col_data_map[col] = input_df[col]

    # Step 2: Add mapped columns (e.g., Style Code → productId)
    for input_col in input_df.columns:
        matches = mapping_df[mapping_df.iloc[:, 1] == input_col]  # Column B = Input Header
        for _, row in matches.iterrows():
            mapped_col = row[3]  # Column D = Output Column
            if mapped_col not in col_data_map:
                col_data_map[mapped_col] = input_df[input_col]

    # Create final DataFrame for output
    final_df = pd.DataFrame(col_data_map)

    # Write to Values tab
    for j, col_name in enumerate(final_df.columns, start=1):
        ws_values.cell(row=1, column=j, value=col_name)
    for i, row in enumerate(final_df.itertuples(index=False), start=2):
        for j, value in enumerate(row, start=1):
            ws_values.cell(row=i, column=j, value=value)

    # Write to Types tab (Row 1–4)
    for col_idx, col_name in enumerate(final_df.columns, start=2):
        ws_type.cell(row=1, column=col_idx, value=col_name)
        ws_type.cell(row=2, column=col_idx, value=col_name)

        match_row = mapping_df[mapping_df.iloc[:, 3] == col_name]  # Column D = Output Column
        if not match_row.empty:
            ws_type.cell(row=3, column=col_idx, value=match_row.iloc[0, 3])  # Row 3: Field
            ws_type.cell(row=4, column=col_idx, value=match_row.iloc[0, 4])  # Row 4: Type
        else:
            ws_type.cell(row=3, column=col_idx, value="Not Found")
            ws_type.cell(row=4, column=col_idx, value="Not Found")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def process_auto_mapping_mode(input_file):
    input_df = pd.read_excel(input_file)
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_values = wb["Values"]
    ws_type = wb["Types"]

    # Values tab
    for j, col_name in enumerate(input_df.columns, start=1):
        ws_values.cell(row=1, column=j, value=col_name)
    for i, row in enumerate(input_df.itertuples(index=False), start=2):
        for j, value in enumerate(row, start=1):
            ws_values.cell(row=i, column=j, value=value)

    # Type tab Row 1 & 2
    for col_idx, header in enumerate(input_df.columns, start=2):
        ws_type.cell(row=1, column=col_idx, value=header)
        ws_type.cell(row=2, column=col_idx, value=header)

    # Type tab Row 3 & Row 4 auto-filled
    for col_idx, header in enumerate(input_df.columns, start=2):
        if ws_type.cell(row=2, column=col_idx).value:
            ws_type.cell(row=3, column=col_idx, value="mandatory")
            if "image" in str(ws_type.cell(row=2, column=col_idx).value).lower():
                ws_type.cell(row=4, column=col_idx, value="imageurlarray")
            else:
                ws_type.cell(row=4, column=col_idx, value="string")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("📊 SKU Template Automation Tool")

mode = st.selectbox("Select Mode", ["Mapping", "Auto-Mapping"])
input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if input_file:
    if mode == "Mapping":
        mapping_df = load_mapping()
        if st.button("Generate Output (Mapping)"):
            with st.spinner("Processing with Mapping..."):
                result = process_mapping_mode(input_file, mapping_df)
                st.success("✅ Output Generated!")
                st.download_button("📥 Download Output", data=result, file_name="output_template.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        if st.button("Generate Output (Auto-Mapping)"):
            with st.spinner("Processing with Auto-Mapping..."):
                result = process_auto_mapping_mode(input_file)
                st.success("✅ Output Generated!")
                st.download_button("📥 Download Output", data=result, file_name="output_template.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.caption("Built for Rubick.ai | By Vishnu Sai")
