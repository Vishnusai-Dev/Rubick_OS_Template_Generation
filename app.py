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

# substrings used to find worksheets in the mapping file
MAPPING_SHEET_KEY = "mapping"
CLIENT_SHEET_KEY  = "mappedclientname"

# ╭───────────────── NORMALISERS & HELPERS ─────────────────╮
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return "".join(str(s).split()).lower()

def clean_header(header: str) -> str:
    return str(header).replace(".", " ").strip()

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

# ╭──────────── SIZE & COLOR COLUMN IDENTIFICATION ───────────╮
SIZE_TOKENS = {
    "XS","S","M","L","XL","XXL","XXXL","2XL","3XL","4XL","5XL",
    "6","8","10","12","14","16","18","20","22","24","26","28","30","32","34","36","38","40","42","44","46","48","50","52"
}
COLOR_TOKENS = {
    "BLACK","WHITE","RED","BLUE","GREEN","YELLOW","ORANGE","PINK","PURPLE","VIOLET",
    "BROWN","BEIGE","MAROON","NAVY","CYAN","TEAL","GREY","GRAY","MAGENTA","GOLD","SILVER",
    "MULTI","MULTICOLOUR","MULTICOLOR","CREAM","OFF WHITE","OFFWHITE","IVORY","OLIVE",
    "LAVENDER","PEACH","TURQUOISE","MUSTARD","KHAKI"
}

def _ratio_match(series: pd.Series, token_set, extra_check=None) -> float:
    if series is None or series.empty:
        return 0.0
    s = series.dropna().astype(str)
    if s.empty:
        return 0.0
    def hit(x: str) -> bool:
        v = x.strip().upper()
        if v in token_set:
            return True
        if extra_check:
            return extra_check(v)
        return False
    return s.apply(hit).mean()

def _looks_size_value(v: str) -> bool:
    # Numeric sizes like 26, 32, 42 (2-digit), or 3-digit like 100
    return bool(re.fullmatch(r"\d{2,3}", v))

def _looks_color_value(v: str) -> bool:
    # allow things like "Light Blue", "Dark-Grey"
    v2 = re.sub(r"[^\w ]"," ", v).strip()
    parts = [p for p in v2.split() if p]
    return any(p in COLOR_TOKENS for p in (p.upper() for p in parts))

def find_size_column(df: pd.DataFrame) -> str | None:
    # header-led first
    for c in df.columns:
        if "size" in norm(c):
            return c
    # fallback: value-led (pick best ratio)
    best_col, best_ratio = None, 0.0
    for c in df.columns:
        r = _ratio_match(df[c], SIZE_TOKENS, _looks_size_value)
        if r > best_ratio:
            best_col, best_ratio = c, r
    return best_col if best_ratio >= 0.35 else None

def find_color_column(df: pd.DataFrame) -> str | None:
    # header-led first
    for c in df.columns:
        nc = norm(c)
        if "color" in nc or "colour" in nc:
            return c
    # fallback: value-led (pick best ratio)
    best_col, best_ratio = None, 0.0
    for c in df.columns:
        r = _ratio_match(df[c], COLOR_TOKENS, _looks_color_value)
        if r > best_ratio:
            best_col, best_ratio = c, r
    return best_col if best_ratio >= 0.35 else None
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

def _select_sheet_and_rows(xl: pd.ExcelFile, template_type: str):
    """
    Returns (sheet_name, header_row_1based, data_start_row_1based)
    """
    if template_type == "Amazon":
        if "Template" not in xl.sheet_names:
            return None, None, None
        return "Template", 2, 4
    if template_type == "Flipkart":
        if len(xl.sheet_names) < 3:
            return None, None, None
        return xl.sheet_names[2], 1, 5
    if template_type == "Myntra":
        if len(xl.sheet_names) < 2:
            return None, None, None
        return xl.sheet_names[1], 3, 4
    if template_type == "Ajio":
        if len(xl.sheet_names) < 3:
            return None, None, None
        return xl.sheet_names[2], 2, 3
    if template_type == "TataCliq":
        if len(xl.sheet_names) < 1:
            return None, None, None
        return xl.sheet_names[0], 4, 6
    # General
    return xl.sheet_names[0], 1, 2  # current process: header row 1, data row 2

def _rebase_header(df_raw: pd.DataFrame, header_row_1b: int, data_start_row_1b: int) -> pd.DataFrame:
    """
    Make the chosen header row become df.columns, and data start from data_start_row_1b.
    """
    if df_raw is None or df_raw.empty:
        return df_raw
    hdr_idx = max(header_row_1b - 1, 0)
    data_idx = max(data_start_row_1b - 1, 0)

    # Create headers
    headers = df_raw.iloc[hdr_idx].astype(str).tolist()
    # Deduplicate headers to avoid collisions
    seen = {}
    dedup_headers = []
    for h in headers:
        h0 = h if h and h.lower() != "nan" else ""
        if h0 not in seen:
            seen[h0] = 1
            dedup_headers.append(h0)
        else:
            seen[h0] += 1
            dedup_headers.append(f"{h0}_{seen[h0]}")

    df = df_raw.iloc[data_idx:].copy()
    df.columns = dedup_headers
    # Drop fully empty rows/cols
    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.to_series().apply(lambda c: df[c].isna().all())]
    df.reset_index(drop=True, inplace=True)
    return df

def read_input_excel(uploaded_file, template_type: str):
    """
    Read Excel with rules depending on template_type.
    Works for XLSX, XLSM, XLS (reads from an in-memory buffer).
    # If this template comes from one of the listed marketplaces, select the right type from the dropdown.
    """
    # Important: read into an in-memory buffer to avoid engine/seek issues
    file_bytes = uploaded_file.read()
    excel_buffer = BytesIO(file_bytes)
    xl = pd.ExcelFile(excel_buffer)  # pandas auto-detects engine

    sheet_name, hdr_1b, data_1b = _select_sheet_and_rows(xl, template_type)
    if sheet_name is None:
        st.warning("❌ Required sheet not found or insufficient sheets for the selected type.")
        return None

    # Read the whole sheet with no header, then rebase
    df_raw = xl.parse(sheet_name, header=None)
    if df_raw is None or df_raw.empty:
        st.warning("❌ Selected sheet is empty.")
        return None

    try:
        df = _rebase_header(df_raw, hdr_1b, data_1b)
    except Exception as e:
        st.warning(f"❌ Failed to rebase header/data rows: {e}")
        return None

    return df

def process_file(uploaded_file, mode: str, template_type: str, mapping_df: pd.DataFrame | None = None):
    src_df = read_input_excel(uploaded_file, template_type)
    if src_df is None:
        return None  # stop processing if invalid sheet

    # ────────── BUILD columns_meta ──────────
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
                if str(row[DUP_KEY]).strip().lower().startswith("yes"):
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

    # ────────── IDENTIFY SIZE (Option 1) & COLOR (Option 2) COLUMNS ──────────
    size_col  = find_size_column(src_df)
    color_col = find_color_column(src_df)
    option1_data = src_df[size_col]  if size_col  else pd.Series([""]*len(src_df))
    option2_data = src_df[color_col] if color_col else pd.Series([""]*len(src_df))

    # ────────── BUILD THE WORKBOOK ──────────
    try:
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
    except Exception as e:
        st.error(f"❌ Could not load template: {e}")
        return None

    if "Values" not in wb.sheetnames or "Types" not in wb.sheetnames:
        st.error("❌ Template must contain 'Values' and 'Types' sheets.")
        return None

    ws_vals  = wb["Values"]
    ws_types = wb["Types"]

    # Clear existing content (optional safety)
    for ws in (ws_vals, ws_types):
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.value = None

    # Write main mapped/auto-mapped columns to Values and Types
    for j, meta in enumerate(columns_meta, start=1):
        header_display = clean_header(meta["out"])
        ws_vals.cell(row=1, column=j, value=header_display)
        col_series = src_df[meta["src"]] if meta["src"] in src_df.columns else pd.Series([None]*len(src_df))
        for i, v in enumerate(col_series.tolist(), start=2):
            cell = ws_vals.cell(row=i, column=j)
            if pd.isna(v) or v == "":
                cell.value = None
            else:
                if str(meta["row4"]).lower() in ("string", "imageurlarray"):
                    cell.value = str(v)
                    cell.number_format = "@"
                else:
                    cell.value = v

        # Types sheet (shift by +2 columns as per your template logic)
        tcol = j + 2
        ws_types.cell(row=1, column=tcol, value=header_display)
        ws_types.cell(row=2, column=tcol, value=header_display)
        ws_types.cell(row=3, column=tcol, value=meta["row3"])
        ws_types.cell(row=4, column=tcol, value=meta["row4"])

    # ────────── APPEND OPTION 1 & OPTION 2 TO VALUES ──────────
    opt1_col = len(columns_meta) + 1
    opt2_col = len(columns_meta) + 2
    if size_col:
        ws_vals.cell(row=1, column=opt1_col, value="Option 1")
        for i, v in enumerate(option1_data.tolist(), start=2):
            ws_vals.cell(row=i, column=opt1_col, value=(None if pd.isna(v) or str(v).strip()=="" else v))
    if color_col:
        ws_vals.cell(row=1, column=opt2_col, value="Option 2")
        for i, v in enumerate(option2_data.tolist(), start=2):
            ws_vals.cell(row=i, column=opt2_col, value=(None if pd.isna(v) or str(v).strip()=="" else v))

    # ────────── APPEND OPTION 1 & OPTION 2 TO TYPES ──────────
    if size_col:
        t1_col = opt1_col + 2
        ws_types.cell(row=1, column=t1_col, value="Option 1")
        ws_types.cell(row=2, column=t1_col, value="Option 1")
        ws_types.cell(row=3, column=t1_col, value="non mandatory")
        ws_types.cell(row=4, column=t1_col, value="string")
        unique_opt1 = pd.Series([str(x) for x in option1_data.dropna().astype(str).tolist() if str(x).strip()])
        for i, v in enumerate(sorted(unique_opt1.unique().tolist()), start=5):
            ws_types.cell(row=i, column=t1_col, value=v)

    if color_col:
        t2_col = opt2_col + 2
        ws_types.cell(row=1, column=t2_col, value="Option 2")
        ws_types.cell(row=2, column=t2_col, value="Option 2")
        ws_types.cell(row=3, column=t2_col, value="non mandatory")
        ws_types.cell(row=4, column=t2_col, value="string")
        unique_opt2 = pd.Series([str(x) for x in option2_data.dropna().astype(str).tolist() if str(x).strip()])
        for i, v in enumerate(sorted(unique_opt2.unique().tolist()), start=5):
            ws_types.cell(row=i, column=t2_col, value=v)

    # Done
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
template_type = st.selectbox(
    "Select Type of Input",
    ["General","Amazon","Flipkart","Myntra","Ajio","TataCliq"],
    help="If this template comes from one of the listed marketplaces, select the right one."
)

input_file = st.file_uploader(
    "Upload Input Excel File (.xlsx, .xls, .xlsm)",
    type=["xlsx", "xls", "xlsm"]
)

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
