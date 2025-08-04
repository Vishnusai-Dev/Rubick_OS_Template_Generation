import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH = "Mapping - Automation.xlsx"

@st.cache_data
def load_mapping():
    return pd.read_excel(MAPPING_PATH)

def process_mapping_mode(input_file, mapping_df):
    input_df = pd.read_excel(input_file)

    # Create output DataFrame
    output_df = input_df.copy()

    # Track added columns
    added_columns = []

    for _, row in mapping_df.iterrows():
        input_col = row[1]  # Column B: Input Header
        output_col = row[2]  # Column C: Output Header
        if input_col in input_df.columns:
            output_df[output_col] = input_df[input_col]
            added_columns.append((output_col, row[3], row[4]))  # Output Header, Level, DataType

    # Prepare Excel
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_values = wb.create_sheet("Values")
    ws_type = wb.create_sheet("Types")

    # Values sheet
    for j, col in enumerate(output_df.columns, start=1):
        ws_values.cell(row=1, column=j, value=col)
    for i, row in enumerate(output_df.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            ws_values.cell(row=i, column=j, value=val)

    # Types sheet
    for col_idx, col in enumerate(output_df.columns, start=2):
        ws_type.cell(row=1, column=col_idx, value=col)
        ws_type.cell(row=2, column=col_idx, value=col)

        match = mapping_df[mapping_df.iloc[:, 2] == col]
        if not match.empty:
            ws_type.cell(row=3, column=col_idx, value=match.iloc[0, 3])  # Level
            ws_type.cell(row=4, column=col_idx, value=match.iloc[0, 4])  # DataType
        else:
            ws_type.cell(row=3, column=col_idx, value="Not Found")
            ws_type.cell(row=4, column=col_idx, value="Not Found")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def process_auto_mapping_mode(input_file):
    input_df = pd.read_excel(input_file)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_values = wb.create_sheet("Values")
    ws_type = wb.create_sheet("Types")

    # Values tab
    for j, col in enumerate(input_df.columns, start=1):
        ws_values.cell(row=1, column=j, value=col)
    for i, row in enumerate(input_df.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            ws_values.cell(row=i, column=j, value=val)

    # Types tab
    for col_idx, col in enumerate(input_df.columns, start=2):
        ws_type.cell(row=1, column=col_idx, value=col)
        ws_type.cell(row=2, column=col_idx, value=col)
        ws_type.cell(row=3, column=col_idx, value="mandatory")
        if "image" in col.lower():
            ws_type.cell(row=4, column=col_idx, value="imageurlarray")
        else:
            ws_type.cell(row=4, column=col_idx, value="string")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# UI
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
