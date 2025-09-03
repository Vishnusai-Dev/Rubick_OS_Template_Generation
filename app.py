import streamlit as st
import pandas as pd
import openpyxl
import os
from io import BytesIO

# ──────────────── CONSTANTS ────────────────
TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH = "Mapping - Automation.xlsx"
ATTR_KEY = "attributes"
TARGET_KEY = "fieldname"

# Predefined size and color lists
SIZE_LIST = ["xxs", "xs", "s", "m", "l", "xl", "xxl", "xxxl", "free size", "28", "30", "32", "34", "36", "38", "40"]
COLOR_LIST = ["red", "blue", "green", "black", "white", "yellow", "pink", "grey", "gray", "purple", "orange", "brown"]

# ──────────────── HELPERS ────────────────
def read_excel_auto(uploaded_file):
    """Read Excel files (xls, xlsx, xlsm) with correct engine."""
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    data = uploaded_file.read()
    buffer = BytesIO(data)
    if ext == ".xls":
        engine = "xlrd"  # ensure xlrd==1.2.0
    else:
        engine = "openpyxl"
    return pd.ExcelFile(buffer, engine=engine)

def load_mapping():
    """Load mapping from default mapping file."""
    if not os.path.exists(MAPPING_PATH):
        st.warning("Mapping file not found. Using empty mapping.")
        return pd.DataFrame(columns=[ATTR_KEY, TARGET_KEY])
    mapping_df = pd.read_excel(MAPPING_PATH)
    mapping_df.columns = mapping_df.columns.str.strip().str.lower()
    return mapping_df

def detect_size_color(value):
    """Detect if value is size or color from predefined lists."""
    val = str(value).strip().lower()
    if val in SIZE_LIST:
        return "size"
    if val in COLOR_LIST:
        return "color"
    return None

def match_and_transform(input_df, mapping_df):
    """
    Row-wise matching:
    - Option1 always size (from size list if matched)
    - Option2 always color (from color list if matched)
    - Other mapped fields preserved
    """
    input_df.columns = input_df.columns.str.strip().str.lower()
    mapping_df[ATTR_KEY] = mapping_df[ATTR_KEY].astype(str).str.strip().str.lower()
    mapping_df[TARGET_KEY] = mapping_df[TARGET_KEY].astype(str).str.strip().str.lower()
    output_df = input_df.copy()

    # Ensure option columns
    if "option1" not in output_df.columns:
        output_df["option1"] = ""
    if "option2" not in output_df.columns:
        output_df["option2"] = ""

    # Row-wise detection
    for idx, row in input_df.iterrows():
        for col in input_df.columns:
            value = str(row[col]).strip()
            if not value or value.lower() == "nan":
                continue

            # Check predefined lists
            detected = detect_size_color(value)
            if detected == "size":
                output_df.at[idx, "option1"] = value
            elif detected == "color":
                output_df.at[idx, "option2"] = value
            else:
                # If not in size/color lists, check mapping
                attr = col.strip().lower()
                mapped = mapping_df[mapping_df[ATTR_KEY] == attr]
                if not mapped.empty:
                    target = mapped.iloc[0][TARGET_KEY]
                    if "size" in target:
                        output_df.at[idx, "option1"] = value
                    elif "color" in target or "colour" in target:
                        output_df.at[idx, "option2"] = value
                    else:
                        output_df.at[idx, col] = value
    return output_df

def fill_template(input_df, mapping_df):
    """Fill template with processed Values + Type tabs."""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    values_ws = wb["Values"]
    type_ws = wb["Type"]

    # Clear both sheets
    for ws in [values_ws, type_ws]:
        for row in ws["A1:Z500"]:
            for cell in row:
                cell.value = None

    # Values tab: write headers + rows
    for c_idx, col_name in enumerate(input_df.columns, start=1):
        values_ws.cell(row=1, column=c_idx).value = col_name
    for r_idx, row in enumerate(input_df.values.tolist(), start=2):
        for c_idx, val in enumerate(row, start=1):
            values_ws.cell(row=r_idx, column=c_idx).value = val

    # Type tab: row1+2 headers, row3+4 mapping
    headers = input_df.columns.tolist()
    for c_idx, col_name in enumerate(headers, start=1):
        type_ws.cell(row=1, column=c_idx).value = col_name
        type_ws.cell(row=2, column=c_idx).value = col_name
        mapped = mapping_df[mapping_df[ATTR_KEY] == col_name.lower()]
        if not mapped.empty:
            type_ws.cell(row=3, column=c_idx).value = mapped.iloc[0][TARGET_KEY]
            type_ws.cell(row=4, column=c_idx).value = mapped.iloc[0][TARGET_KEY]

    # Save to buffer
    output_buffer = BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

# ──────────────── STREAMLIT APP ────────────────
st.set_page_config(page_title="Template Mapper", layout="wide")
st.title("Rubick OS – Excel Template Automation")

# Mode: Mapping vs Direct
mode = st.radio("Select Mode", ["Mapping", "Direct"])

# Template Type Dropdown
template_type = st.selectbox(
    "Select Template Type",
    ["General", "Amazon", "Flipkart", "Myntra", "Ajio", "TataCliq"]
)

uploaded_file = st.file_uploader(
    "Upload your input Excel (xls/xlsx/xlsm)",
    type=["xls", "xlsx", "xlsm"]
)

if uploaded_file:
    st.info("Processing your file...")
    mapping_df = load_mapping()

    # Allow user to edit mapping if in Mapping mode
    if mode == "Mapping" and not mapping_df.empty:
        st.subheader("Edit Mapping Table")
        mapping_df = st.data_editor(mapping_df, num_rows="dynamic")

    # Read input Excel
    try:
        input_xl = read_excel_auto(uploaded_file)
        sheet_name = input_xl.sheet_names[0]
        input_df = pd.read_excel(input_xl, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        st.stop()

    # Apply mapping and size/color detection
    processed_df = match_and_transform(input_df, mapping_df)

    # Fill template with processed data
    output_buffer = fill_template(processed_df, mapping_df)

    # Download button
    st.success("Template generated successfully!")
    st.download_button(
        label="Download Processed Template",
        data=output_buffer,
        file_name="sku_template_filled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
