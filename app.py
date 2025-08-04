import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ───────────── File paths ─────────────
TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH  = "Mapping - Automation.xlsx"

# keys we expect – all lower-case, no spaces
ATTR_KEY   = "attributes"
TARGET_KEY = "fieldname"
MAND_KEY   = "mandatoryornot"
TYPE_KEY   = "fieldtype"
DUP_KEY    = "duplicatestobecreated"

MAPPING_SHEET_KEY = "mapping"            # part of the name is fine
CLIENT_SHEET_KEY  = "mappedclientname"   # part of the name is fine
# ───────────────────────────────────────


# ╭───────────────── helper: normalise text ─────────────────╮
def norm(s) -> str:
    """Trim, lower-case, collapse internal whitespace → key string."""
    if pd.isna(s):
        return ""
    return " ".join(str(s).strip().split()).lower()
# ╰───────────────────────────────────────────────────────────╯


@st.cache_data
def load_mapping():
    """Return (mapping_df, client_names) using tolerant name matching."""
    xl = pd.ExcelFile(MAPPING_PATH)

    # ── choose mapping sheet ──
    map_sheet = next(
        (s for s in xl.sheet_names if MAPPING_SHEET_KEY in norm(s)),
        xl.sheet_names[0]
    )
    mapping_df = xl.parse(map_sheet)
    mapping_df.rename(columns={c: norm(c) for c in mapping_df.columns},
                      inplace=True)         # normalise headers

    # ── choose client sheet ──
    client_names = []
    client_sheet = next((s for s in xl.sheet_names if CLIENT_SHEET_KEY in norm(s)),
                        None)
    if client_sheet:
        raw = xl.parse(client_sheet, header=None)
        client_names = [str(x).strip() for x in raw.values.flatten()
                        if pd.notna(x) and str(x).strip()]

    return mapping_df, client_names


def process_file(input_file, mode: str, mapping_df: pd.DataFrame | None = None):
    """Return BytesIO workbook for given mode."""
    src_df = pd.read_excel(input_file)
    columns_meta = []

    if mode == "Mapping" and mapping_df is not None:
        for col in src_df.columns:
            matches = mapping_df[mapping_df[ATTR_KEY] == norm(col)]

            # keep the original column
            if not matches.empty:
                row3 = matches.iloc[0][MAND_KEY]
                row4 = matches.iloc[0][TYPE_KEY]
            else:
                row3 = row4 = "Not Found"

            columns_meta.append({"src": col, "out": col, "row3": row3, "row4": row4})

            # duplicates flagged yes
            for _, row in matches.iterrows():
                if str(row[DUP_KEY]).lower().startswith("yes"):
                    new_hdr = row[TARGET_KEY] if pd.notna(row[TARGET_KEY]) else col
                    if new_hdr != col:
                        columns_meta.append({
                            "src": col, "out": new_hdr,
                            "row3": row[MAND_KEY],
                            "row4": row[TYPE_KEY]
                        })
    else:   # Auto-Mapping
        for col in src_df.columns:
            dtype = "imageurlarray" if "image" in norm(col) else "string"
            columns_meta.append({"src": col, "out": col,
                                 "row3": "mandatory", "row4": dtype})

    # ── build workbook ──
    wb        = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals   = wb["Values"]
    ws_types  = wb["Types"]

    for j, m in enumerate(columns_meta, start=1):
        ws_vals.cell(row=1, column=j, value=m["out"])
        for i, v in enumerate(src_df[m["src"]].tolist(), start=2):
            ws_vals.cell(row=i, column=j, value=v)

    for j, m in enumerate(columns_meta, start=2):
        ws_types.cell(row=1, column=j, value=m["out"])
        ws_types.cell(row=2, column=j, value=m["out"])
        ws_types.cell(row=3, column=j, value=m["row3"])
        ws_types.cell(row=4, column=j, value=m["row4"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ──────────────── Streamlit UI ────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("📊 SKU Template Automation Tool")

mapping_df, client_names = load_mapping()

if client_names:
    st.info("🗂️  **Mapped clients available:** " + ", ".join(client_names))
else:
    st.warning("⚠️  No client list found / sheet name mismatch.")

mode       = st.selectbox("Select Mode", ["Mapping", "Auto-Mapping"])
input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if input_file and st.button(f"Generate Output ({mode})"):
    with st.spinner("Processing…"):
        result = process_file(input_file, mode,
                              mapping_df if mode == "Mapping" else None)
        st.success("✅ Output Generated!")
        st.download_button("📥 Download Output",
                           data=result,
                           file_name="output_template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.caption("Built for Rubick.ai | By Vishnu Sai")
