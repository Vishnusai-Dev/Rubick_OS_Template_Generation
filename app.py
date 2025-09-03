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
# ──────────────────────────────────────────────────────────────

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

def process_file(input_file, mode: str, mapping_df: pd.DataFrame | None = None):
    src_df = pd.read_excel(input_file)

    # ────────── DROP COMPLETELY EMPTY COLUMNS ──────────
    src_df.dropna(axis=1, how='all', inplace=True)

    columns_meta = []

    # ────────── BUILD columns_meta ──────────
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
    else:  # Auto-Mapping
        for col in src_df.columns:
            dtype = "imageurlarray" if is_image_column(norm(col), src_df[col]) else "string"
            columns_meta.append({"src": col, "out": col, "row3": "mandatory", "row4": dtype})

    # ────────── ADD OPTION 1 & OPTION 2 ──────────
    size_values = {"XS","S","M","L","XL","XXL","2XL","3XL","XXXL"}
    color_values = {
        "RED","WHITE","GREEN","BLUE","YELLOW","BLACK","BROWN",
        "ORANGE","PURPLE","PINK","GREY","GRAY","BEIGE","MAROON","NAVY"
    }
    def exact_match(val, valid_set):
        if pd.isna(val):
            return ""
        s = str(val).strip().upper()
        return s if s in valid_set else ""

    option1_data = pd.Series([""]*len(src_df), dtype=str)
    option2_data = pd.Series([""]*len(src_df), dtype=str)

    # collect matches from ALL size/color columns without overwriting
    for col in src_df.columns:
        if "size" in norm(col):
            temp = src_df[col].apply(lambda x: exact_match(x, size_values))
            option1_data = option1_data.mask(option1_data=="", temp)
        if "color" in norm(col) or "colour" in norm(col):
            temp = src_df[col].apply(lambda x: exact_match(x, color_values))
            option2_data = option2_data.mask(option2_data=="", temp)

    # ────────── BUILD THE WORKBOOK ──────────
    wb        = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals   = wb["Values"]
    ws_types  = wb["Types"]

    # Write main mapped/auto-mapped columns to Values and Types
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

    # ────────── APPEND OPTION 1 & OPTION 2 TO VALUES ──────────
    opt1_col = len(columns_meta) + 1
    opt2_col = len(columns_meta) + 2
    ws_vals.cell(row=1, column=opt1_col, value="Option 1")
    ws_vals.cell(row=1, column=opt2_col, value="Option 2")
    for i, v in enumerate(option1_data.tolist(), start=2):
        ws_vals.cell(row=i, column=opt1_col, value=v if v else None)
    for i, v in enumerate(option2_data.tolist(), start=2):
        ws_vals.cell(row=i, column=opt2_col, value=v if v else None)

    # ────────── APPEND OPTION 1 & OPTION 2 TO TYPES ──────────
    t1_col = opt1_col + 2
    t2_col = opt2_col + 2
    ws_types.cell(row=1, column=t1_col, value="Option 1")
    ws_types.cell(row=2, column=t1_col, value="Option 1")
    ws_types.cell(row=3, column=t1_col, value="non mandatory")
    ws_types.cell(row=4, column=t1_col, value="select")
    ws_types.cell(row=1, column=t2_col, value="Option 2")
    ws_types.cell(row=2, column=t2_col, value="Option 2")
    ws_types.cell(row=3, column=t2_col, value="non mandatory")
    ws_types.cell(row=4, column=t2_col, value="select")
    unique_opt1 = pd.Series([x for x in option1_data.unique() if x])
    unique_opt2 = pd.Series([x for x in option2_data.unique() if x])
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

mode       = st.selectbox("Select Mode", ["Mapping", "Auto-Mapping"])
input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])
if input_file and st.button(f"Generate Output ({mode})"):
    with st.spinner("Processing…"):
        result = process_file(input_file, mode, mapping_df if mode == "Mapping" else None)
        st.success("✅ Output Generated!")
        st.download_button(
            "📥 Download Output",
            data=result,
            file_name="output_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("---")
st.caption("Built for Rubick.ai | By Vishnu Sai")
