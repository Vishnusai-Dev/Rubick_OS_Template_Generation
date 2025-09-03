import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ───────────────────────── FILE PATHS ─────────────────────────
TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH  = "Mapping - Automation.xlsx"

# ─────────────────── INTERNAL COLUMN KEYS ───────────────────
ATTR_KEY   = "attributes"
TARGET_KEY = "fieldname"

# ──────────────────── HELPER FUNCTIONS ────────────────────
def load_mapping():
    df = pd.read_excel(MAPPING_PATH)
    mapping = {}
    for _, row in df.iterrows():
        attr = str(row[ATTR_KEY]).strip().lower()
        target = str(row[TARGET_KEY]).strip()
        mapping[attr] = target
    return mapping

def detect_size(text):
    # Common patterns: S, M, L, XL, XXL, 32, 34, 36, etc.
    return bool(re.search(r"\b(XXL|XL|L|M|S|XS|\d{2})\b", str(text), re.IGNORECASE))

def detect_color(text):
    # Simplified color check
    common_colors = ["red", "blue", "green", "black", "white", "yellow", "pink", "purple", "orange", "grey", "gray", "brown"]
    return any(color in str(text).lower() for color in common_colors)

def process_file(input_df, template_df, mapping):
    output_df = template_df.copy()

    # ────────────────────────── VALUES TAB ──────────────────────────
    if "Values" not in output_df.sheetnames:
        st.warning("Template does not contain 'Values' tab.")
        return None
    
    # Write full input into "Values"
    with pd.ExcelWriter(BytesIO(), engine='openpyxl') as writer:
        input_df.to_excel(writer, index=False, sheet_name="Values")

    # ────────────────────────── TYPE TAB ──────────────────────────
    type_sheet = output_df["Type"]
    for idx, col in enumerate(input_df.columns, start=1):
        # Row 1 & 2 from header
        type_sheet.cell(row=1, column=idx).value = col
        type_sheet.cell(row=2, column=idx).value = col

        # Row 3 & 4 using mapping
        attr_lower = col.lower().strip()
        mapped_field = mapping.get(attr_lower, "")
        type_sheet.cell(row=3, column=idx).value = mapped_field
        type_sheet.cell(row=4, column=idx).value = mapped_field

    # ────────────────────────── OPTIONS TAB ──────────────────────────
    if "Options" not in output_df.sheetnames:
        st.warning("Template does not contain 'Options' tab.")
        return None

    options_sheet = output_df["Options"]
    row_idx = 2  # assuming headers at row 1
    for _, row in input_df.iterrows():
        size_val, color_val = "", ""
        for val in row:
            if detect_size(val):
                size_val = val
            elif detect_color(val):
                color_val = val
        
        options_sheet.cell(row=row_idx, column=1).value = size_val  # Option 1 = Size
        options_sheet.cell(row=row_idx, column=2).value = color_val  # Option 2 = Color
        row_idx += 1

    return output_df

# ─────────────────────── STREAMLIT APP ───────────────────────
st.title("Excel Attribute Mapper (Auto Output)")

uploaded_file = st.file_uploader(
    "Upload your Excel file (.xlsx, .xls, .xlsm)",
    type=["xlsx", "xls", "xlsm"]
)

if uploaded_file:
    input_df = pd.read_excel(uploaded_file)
    mapping = load_mapping()
    template_wb = pd.ExcelWriter(BytesIO(), engine='openpyxl')
    template_wb.book = pd.read_excel(TEMPLATE_PATH, sheet_name=None)  # Load as workbook

    output_wb = process_file(input_df, template_wb.book, mapping)
    if output_wb:
        buffer = BytesIO()
        output_wb.save(buffer)
        buffer.seek(0)

        st.success("Processing complete! Click below to download:")
        st.download_button(
            label="Download Processed Excel",
            data=buffer,
            file_name="processed_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Failed to process file. Check logs.")
