import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO

# ───────────────────────── FILE PATHS ─────────────────────────
TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH  = "Mapping - Automation.xlsx"

# ─────────────────── INTERNAL   COLUMN KEYS ───────────────────
ATTR_KEY   = "attributes"
TARGET_KEY = "fieldname"
MAND_KEY   = "mandatoryornot"
TYPE_KEY   = "fieldtype"
DUP_KEY    = "duplicatestobecreated"

# substrings used to find worksheets
MAPPING_SHEET_KEY = "mapping"
CLIENT_SHEET_KEY  = "mappedclientname"

# ╭───────────────── NORMALISERS & HELPERS ─────────────────╮
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return "".join(str(s).split()).lower()

def clean_header(header: str) -> str:
    return header.replace(".", " ").strip()

IMAGE_EXT_RE = re.compile(r"(?i)\.(jpe?g|png|gif|bmp|webp|tiff?)$")
IMAGE_KEYWORDS = {
    "image","img","picture","photo","thumbnail","thumb",
    "hero","front","back","url"
}
def is_image_column(col_header_norm: str, series: pd.Series) -> bool:
    header_hit = any(k in col_header_norm for k in IMAGE_KEYWORDS)
    sample = series.dropna().astype(str).head(20)
    ratio  = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0.0
    return header_hit or ratio >= 0.30
# ╰───────────────────────────────────────────────────────────╯

@st.cache_data
def load_mapping():
    xl = pd.ExcelFile(MAPPING_PATH)
    map_sheet = next((s for s in xl.sheet_names if MAPPING_SHEET_KEY in norm(s)), xl.sheet_names[0])
    mapping_df = xl.parse(map_sheet)
    mapping_df.rename(columns={c: norm(c) for c in mapping_df.columns}, inplace=True)
    mapping_df["__attr_key"] = mapping_df[ATTR_KEY].apply(norm)

    client_names = []
    client_sheet = next((s for s in xl.sheet_names if CLIENT_SHEET_KEY in norm(s)), None)
    if client_sheet:
        raw = xl.parse(client_sheet, header=None)
        client_names = [str(x).strip() for x in raw.values.flatten() if pd.notna(x) and str(x).strip()]

    return mapping_df, client_names

def read_input_excel(input_file, template_type: str):
    """Read Excel with rules depending on template_type.
       Supports XLSX, XLSM, and XLS formats.
    """
    xl = pd.ExcelFile(input_file)  # pandas auto-detects type

    if template_type == "Amazon":
        if "Template" not in xl.sheet_names:
            st.warning("❌ 'Template' sheet not found in file.")
            return None
        df = xl.parse("Template", header=1, skiprows=[0,1])  
    elif template_type == "Flipkart":
        if len(xl.sheet_names) < 3:
            st.warning("❌ Flipkart template requires at least 3 sheets.")
            return None
        df = xl.parse(xl.sheet_names[2], header=0, skiprows=[1,2,3,4])  
    elif template_type == "Myntra":
        if len(xl.sheet_names) < 2:
            st.warning("❌ Myntra template requires at least 2 sheets.")
            return None
        df = xl.parse(xl.sheet_names[1], header=2)  
    elif template_type == "Ajio":
        if len(xl.sheet_names) < 3:
            st.warning("❌ Ajio template requires at least 3 sheets.")
            return None
        df = xl.parse(xl.sheet_names[2], header=1)  
    elif template_type == "TataCliq":
        if len(xl.sheet_names) < 1:
            st.warning("❌ TataCliq template requires at least 1 sheet.")
            return None
        df = xl.parse(xl.sheet_names[0], header=3)  
    else:  # General
        df = xl.parse(xl.sheet_names[0], header=0)  

    return df

def process_file(input_file, mode: str, template_type: str, mapping_df: pd.DataFrame | None = None):
    src_df = read_input_excel(input_file, template_type)
    if src_df is None:
        return None  # stop processing if invalid sheet

    columns_meta = []

    if mode == "Mapping" and mapping_df is not None:
        for col in src_df.columns:
            col_key = norm(col)
            matches = mapping_df[mapping_df["__attr_key"] == col_key]
            if not matches.empty:
                row3 = matches.iloc[0][MAND_KEY]
                row4 = matches.iloc[0][TYPE_KEY]
            else:
                row3 = row4 = "Not Found"
            columns_meta.append({"src": col, "out": col, "row3": row3, "row4": row4})
            for _, row in matches.iterrows():
                if str(row[DUP_KEY]).lower().startswith("yes"):
                    new_header = row[TARGET_KEY] if pd.notna(row[TARGET_KEY]) else col
                    if new_header != col:
                        columns_meta.append({
                            "src": col,
                            "out": new_header,
                            "row3": row[MAND_KEY],
                            "row4": row[TYPE_KEY]
                        })
    else:  
        for col in src_df.columns:
            dtype = "imageurlarray" if is_image_column(norm(col), src_df[col]) else "string"
            columns_meta.append({"src": col, "out": col, "row3": "mandatory", "row4": dtype})

    size_col = next((c for c in src_df.columns if "size" in norm(c)), None)
    color_col = next((c for c in src_df.columns if ("color" in norm(c) or "colour" in norm(c))), None)
    option1_data = src_df[size_col] if size_col else pd.Series([""]*len(src_df))
    option2_data = src_df[color_col] if color_col else pd.Series([""]*len(src_df))

    wb        = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals   = wb["Values"]
    ws_types  = wb["Types"]

    for j, meta in enumerate(columns_meta, start=1):
        header_display = clean_header(meta["out"])
        ws_vals.cell(row=1, column=j, value=header_display)
        for i, v in enumerate(src_df[meta["src"]].tolist(), start=2):
            cell = ws_vals.cell(row=i, column=j)
            if pd.isna(v):
                cell.value = None
                continue
            if str(meta["row4"]).lower() in ("string", "imageurlarray"):
                cell.value = str(v)
                cell.number_format = "@"
            else:
                cell.value = v
        tcol = j + 2
        ws_types.cell(row=1, column=tcol, value=header_display)
        ws_types.cell(row=2, column=tcol, value=header_display)
        ws_types.cell(row=3, column=tcol, value=meta["row3"])
        ws_types.cell(row=4, column=tcol, value=meta["row4"])

    opt1_col = len(columns_meta) + 1
    opt2_col = len(columns_meta) + 2
    ws_vals.cell(row=1, column=opt1_col, value="Option 1")
    ws_vals.cell(row=1, column=opt2_col, value="Option 2")
    for i, v in enumerate(option1_data.tolist(), start=2):
        ws_vals.cell(row=i, column=opt1_col, value=v if pd.notna(v) and v != "" else None)
    for i, v in enumerate(option2_data.tolist(), start=2):
        ws_vals.cell(row=i, column=opt2_col, value=v if pd.notna(v) and v != "" else None)

    t1_col = opt1_col + 2
    t2_col = opt2_col + 2
    ws_types.cell(row=1, column=t1_col, value="Option 1")
    ws_types.cell(row=2, column=t1_col, value="Option 1")
    ws_types.cell(row=3, column=t1_col, value="non mandatory")
    ws_types.cell(row=4, column=t1_col, value="string")
    ws_types.cell(row=1, column=t2_col, value="Option 2")
    ws_types.cell(row=2, column=t2_col, value="Option 2")
    ws_types.cell(row=3, column=t2_col, value="non mandatory")
    ws_types.cell(row=4, column=t2_col, value="string")
    unique_opt1 = pd.Series([str(x) for x in option1_data.dropna().unique() if str(x).strip()])
    unique_opt2 = pd.Series([str(x) for x in option2_data.dropna().unique() if str(x).strip()])
    for i, v in enumerate(unique_opt1.tolist(), start=5):
        ws_types.cell(row=i, column=t1_col, value=v)
    for i, v in enumerate(unique_opt2.tolist(), start=5):
        ws_types.cell(row=i, column=t2_col, value=v)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ───────────────────────── STREAMLIT UI ─────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("📊 SKU Template Automation Tool")

mapping_df, client_names = load_mapping()
if client_names:
    st.info("🗂️  **Mapped clients available:** " + ", ".join(client_names))
else:
    st.warning("⚠️  No client list found in the mapping workbook.")

mode          = st.selectbox("Select Mode", ["Mapping", "Auto-Mapping"])
template_type = st.selectbox("Select Type of Input", ["General","Amazon","Flipkart","Myntra","Ajio","TataCliq"])
input_file    = st.file_uploader("Upload Input Excel File", type=["xlsx", "xls", "xlsm"])

if input_file:
    with st.spinner("Processing…"):
        result = process_file(input_file, mode, template_type, mapping_df if mode == "Mapping" else None)
        if result:
            st.success("✅ Output Generated!")
            st.download_button(
                "📥 Download Output",
                data=result,
                file_name="output_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Processing failed. Please check the warnings above.")

st.markdown("---")
st.caption("Built for Rubick.ai | By Vishnu Sai")
