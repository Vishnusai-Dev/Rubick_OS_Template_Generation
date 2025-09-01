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
    """Trim, remove *all* whitespace, lower-case."""
    if pd.isna(s):
        return ""
    return "".join(str(s).split()).lower()

def clean_header(header: str) -> str:
    """Replace each '.' with a single space and trim."""
    return header.replace(".", " ").strip()

# image detection helpers
IMAGE_EXT_RE = re.compile(r"(?i)\.(jpe?g|png|gif|bmp|webp|tiff?)$")
IMAGE_KEYWORDS = {
    "image", "img", "picture", "photo", "thumbnail", "thumb",
    "hero", "front", "back", "url"
}

def is_image_column(col_header_norm: str, series: pd.Series) -> bool:
    """True if header has an image keyword OR ≥30 % of first 20 non-blank values look like image files/URLs."""
    header_hit = any(k in col_header_norm for k in IMAGE_KEYWORDS)
    sample = series.dropna().astype(str).head(20)
    ratio  = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0.0
    return header_hit or ratio >= 0.30
# ╰───────────────────────────────────────────────────────────╯

@st.cache_data
def load_mapping():
    """Return (mapping_df, client_names) with tolerant sheet / header names."""
    xl = pd.ExcelFile(MAPPING_PATH)

    # mapping sheet
    map_sheet = next((s for s in xl.sheet_names if MAPPING_SHEET_KEY in norm(s)),
                     xl.sheet_names[0])
    mapping_df = xl.parse(map_sheet)
    mapping_df.rename(columns={c: norm(c) for c in mapping_df.columns},
                      inplace=True)
    mapping_df["__attr_key"] = mapping_df[ATTR_KEY].apply(norm)

    # client sheet
    client_names = []
    client_sheet = next((s for s in xl.sheet_names if CLIENT_SHEET_KEY in norm(s)),
                        None)
    if client_sheet:
        raw = xl.parse(client_sheet, header=None)
        client_names = [
            str(x).strip() for x in raw.values.flatten()
            if pd.notna(x) and str(x).strip()
        ]

    return mapping_df, client_names

def process_file(input_file, mode: str, mapping_df: pd.DataFrame | None = None):
    """Generate the filled template and return it as BytesIO."""
    src_df = pd.read_excel(input_file)
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
    else:   # Auto-Mapping
        for col in src_df.columns:
            dtype = (
                "imageurlarray"
                if is_image_column(norm(col), src_df[col])
                else "string"
            )
            columns_meta.append({"src": col, "out": col,
                                 "row3": "mandatory", "row4": dtype})

    # ────────── ADD OPTION 1 & OPTION 2 ──────────
    size_values = ["XS","S","M","L","XL","XXL","2XL","3XL","XXXL"]
    color_values = [
        "Red","White","Green","Blue","Yellow","Black","Brown",
        "Orange","Purple","Pink","Grey","Gray","Beige","Maroon","Navy"
    ]

    option1_data = pd.Series(dtype=str)
    option2_data = pd.Series(dtype=str)

    for col in src_df.columns:
        if "size" in norm(col):
            option1_data = pd.concat([option1_data, src_df[col].astype(str)], ignore_index=True)
        if "color" in norm(col) or "colour" in norm(col):
            option2_data = pd.concat([option2_data, src_df[col].astype(str)], ignore_index=True)

    # STRICT MATCH VALIDATION (blank if no match)
    def strict_validate(value, valid_list):
        v = str(value).strip()
        return v if v in valid_list else ""

    option1_data = option1_data.apply(lambda x: strict_validate(x, size_values))
    option2_data = option2_data.apply(lambda x: strict_validate(x, color_values))

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
        ws_vals.cell(row=i, column=opt1_col, value=v if v.strip() else None)
    for i, v in enumerate(option2_data.tolist(), start=2):
        ws_vals.cell(row=i, column=opt2_col, value=v if v.strip() else None)

    # ────────── APPEND OPTION 1 & OPTION 2 TO TYPES ──────────
    t1_col = opt1_col + 2
    t2_col = opt2_col + 2

    ws_types.cell(row=1, column=t1_col, value="Option 1")
    ws_types.cell(row=2, column=t1_col, value="Option 1")
    ws_types.cell(row=3, column=t1_col, value="non mandatory")  # only row3 is marked
    ws_types.cell(row=4, column=t1_col, value="string")

    ws_types.cell(row=1, column=t2_col, value="Option 2")
    ws_types.cell(row=2, column=t2_col, value="Option 2")
    ws_types.cell(row=3, column=t2_col, value="non mandatory")  # only row3 is marked
    ws_types.cell(row=4, column=t2_col, value="string")

    # Unique values from Option1 and Option2 into Types (starting row 5)
    unique_opt1 = pd.Series(option1_data.dropna().unique())
    unique_opt2 = pd.Series(option2_data.dropna().unique())

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
        result = process_file(input_file, mode,
                              mapping_df if mode == "Mapping" else None)
        st.success("✅ Output Generated!")
        st.download_button(
            "📥 Download Output",
            data=result,
            file_name="output_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("---")
st.caption("Built for Rubick.ai | By Vishnu Sai")
