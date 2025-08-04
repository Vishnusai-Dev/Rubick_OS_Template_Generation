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
    """
    Replace every period with a single space and trim the result.
    'Color.Shade' → 'Color Shade'
    """
    return header.replace(".", " ").strip()


# image detection helpers
IMAGE_EXT_RE = re.compile(r"(?i)\.(jpe?g|png|gif|bmp|webp|tiff?)$")
IMAGE_KEYWORDS = {
    "image", "img", "picture", "photo", "thumbnail", "thumb",
    "hero", "front", "back", "url"
}

def is_image_column(col_header_norm: str, series: pd.Series) -> bool:
    """
    True if either:
      • header contains an image keyword, OR
      • ≥30 % of first 20 non-blank values end with an image extension
    """
    header_hit = any(k in col_header_norm for k in IMAGE_KEYWORDS)

    sample = series.dropna().astype(str).head(20)
    value_hit_ratio = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0.0

    return header_hit or value_hit_ratio >= 0.30
# ╰───────────────────────────────────────────────────────────╯



@st.cache_data
def load_mapping():
    """Return (mapping_df, client_names) with robust sheet/header handling."""
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
        client_names = [str(x).strip() for x in raw.values.flatten()
                        if pd.notna(x) and str(x).strip()]

    return mapping_df, client_names



def process_file(input_file, mode: str, mapping_df: pd.DataFrame | None = None):
    """Generate the filled template and return it as BytesIO."""
    src_df = pd.read_excel(input_file)
    columns_meta = []

    # ────────── MAPPING MODE ──────────
    if mode == "Mapping" and mapping_df is not None:
        for col in src_df.columns:
            col_key = norm(col)
            matches = mapping_df[mapping_df["__attr_key"] == col_key]

            # keep original column
            if not matches.empty:
                row3, row4 = matches.iloc[0][MAND_KEY], matches.iloc[0][TYPE_KEY]
            else:
                row3 = row4 = "Not Found"

            columns_meta.append({"src": col, "out": col, "row3": row3, "row4": row4})

            # duplicates flagged “yes”
            for _, row in matches.iterrows():
                if str(row[DUP_KEY]).lower().startswith("yes"):
                    new_header = row[TARGET_KEY] if pd.notna(row[TARGET_KEY]) else col
                    if new_header != col:
                        columns_meta.append({
                            "src": col, "out": new_header,
                            "row3": row[MAND_KEY],
                            "row4": row[TYPE_KEY]
                        })

    # ────────── AUTO-MAPPING MODE ──────────
    else:
        for col in src_df.columns:
            dtype = (
                "imageurlarray"
                if is_image_column(norm(col), src_df[col])
                else "string"
            )
            columns_meta.append({"src": col, "out": col,
                                 "row3": "mandatory", "row4": dtype})

    # ────────── BUILD THE WORKBOOK ──────────
    wb        = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals   = wb["Values"]
    ws_types  = wb["Types"]

    for j, m in enumerate(columns_meta, start=1):
        header_display = clean_header(m["out"])

        # Values sheet
        ws_vals.cell(row=1, column=j, value=header_display)
        for i, v in enumerate(src_df[m["src"]].tolist(), start=2):
            ws_vals.cell(row=i, column=j, value=v)   # type preserved

        # Types sheet
        ws_types.cell(row=1, column=j, value=header_display)
        ws_types.cell(row=2, column=j, value=header_display)
        ws_types.cell(row=3, column=j, value=m["row3"])
        ws_types.cell(row=4, column=j, value=m["row4"])

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
