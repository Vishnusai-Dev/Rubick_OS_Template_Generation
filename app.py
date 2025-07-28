
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

    # Type tab Row 3 & 4 from mapping
    for col_idx, header in enumerate(input_df.columns, start=2):
        match_row = mapping_df[mapping_df.iloc[:, 1] == header]
        if not match_row.empty:
            row3_val = match_row.iloc[0, 3]  # Column D
            row4_val = match_row.iloc[0, 4]  # Column E
            ws_type.cell(row=3, column=col_idx, value=row3_val)
            ws_type.cell(row=4, column=col_idx, value=row4_val)
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
