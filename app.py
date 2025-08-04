import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ───────────────────────── Files & sheet names ──────────────────────────
TEMPLATE_PATH        = "sku-template (4).xlsx"
MAPPING_PATH         = "Mapping - Automation.xlsx"

MAPPING_SHEET_NAME   = "Mapping"               # <-- exact name of main mapping sheet
CLIENT_SHEET_NAME    = "Mapped Client Name"    # <-- exact name of client list sheet
# ───────────────────────── Column headers in mapping sheet ──────────────
ATTR_COL   = "Attributes"            # header exactly as it appears in the user's file
TARGET_COL = "Field Name"            # header you want in the template
MAND_COL   = "Mandatory OR Not"      # value → Types row-3
TYPE_COL   = "Field Type"            # value → Types row-4
DUP_COL    = "Duplicates to be created"  # “yes” → create duplicate
# ─────────────────────────────────────────────────────────────────────────

# ╭──────────────────────────── Helpers ─────────────────────────────╮
@st.cache_data
def load_mapping():
    """
    Returns (mapping_df, client_names).
    • mapping_df   ← sheet `MAPPING_SHEET_NAME`  (fallback = first sheet).
    • client_names ← sheet `CLIENT_SHEET_NAME`   (fallback = []).
    """
    xl = pd.ExcelFile(MAPPING_PATH)

    # --- mapping sheet ---
    if MAPPING_SHEET_NAME in xl.sheet_names:
        mapping_df = xl.parse(MAPPING_SHEET_NAME)
    else:
        st.warning(f"⚠️ Sheet “{MAPPING_SHEET_NAME}” not found – using first sheet.")
        mapping_df = xl.parse(xl.sheet_names[0])

    # --- client sheet ---
    if CLIENT_SHEET_NAME in xl.sheet_names:
        clients_raw  = xl.parse(CLIENT_SHEET_NAME, header=None)
        client_names = [
            str(x).strip()
            for x in clients_raw.values.flatten()
            if pd.notna(x) and str(x).strip()
        ]
    else:
        st.warning(f"⚠️ Sheet “{CLIENT_SHEET_NAME}” not found – client list empty.")
        client_names = []

    return mapping_df, client_names


def process_file(input_file, mode: str, mapping_df: pd.DataFrame | None = None):
    """Return BytesIO of finished workbook (both modes handled here)."""
    src_df = pd.read_excel(input_file)
    columns_meta = []          # drives both Values & Types sheets

    # ────── Mapping mode ──────
    if mode == "Mapping" and mapping_df is not None:
        for col in src_df.columns:
            matches = mapping_df[mapping_df[ATTR_COL] == col]

            # keep the original column
            if not matches.empty:
                ref = matches.iloc[0]
                row3, row4 = ref[MAND_COL], ref[TYPE_COL]
            else:
                row3 = row4 = "Not Found"

            columns_meta.append({"src": col, "out": col, "row3": row3, "row4": row4})

            # create duplicates flagged “yes”
            for _, row in matches.iterrows():
                if str(row[DUP_COL]).strip().lower().startswith("yes"):
                    new_hdr = row[TARGET_COL] if pd.notna(row[TARGET_COL]) else col
                    if new_hdr != col:   # avoid self-duplicate
                        columns_meta.append({
                            "src": col, "out": new_hdr,
                            "row3": row[MAND_COL],
                            "row4": row[TYPE_COL]
                        })

    # ────── Auto-Mapping mode ──────
    else:
        for col in src_df.columns:
            dtype = "imageurlarray" if "image" in col.lower() else "string"
            columns_meta.append({"src": col, "out": col,
                                 "row3": "mandatory", "row4": dtype})

    # ────── Build the workbook ──────
    wb        = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_values = wb["Values"]
    ws_types  = wb["Types"]

    # Values sheet
    for j, meta in enumerate(columns_meta, start=1):
        ws_values.cell(row=1, column=j, value=meta["out"])
        for i, val in enumerate(src_df[meta["src"]].tolist(), start=2):
            ws_values.cell(row=i, column=j, value=val)

    # Types sheet
    for j, meta in enumerate(columns_meta, start=2):
        ws_types.cell(row=1, column=j, value=meta["out"])
        ws_types.cell(row=2, column=j, value=meta["out"])
        ws_types.cell(row=3, column=j, value=meta["row3"])
        ws_types.cell(row=4, column=j, value=meta["row4"])

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out
# ╰──────────────────────────────────────────────────────────────╯


# ───────────────────────────── Streamlit UI ─────────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("📊 SKU Template Automation Tool")

mapping_df, client_names = load_mapping()

if client_names:
    st.info("🗂️  **Mapped clients available:** " + ", ".join(client_names))
else:
    st.warning("⚠️  No client list found in the mapping workbook.")

mode       = st.selectbox("Select Mode", ["Mapping", "Auto-Mapping"])
input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if input_file:
    if st.button(f"Generate Output ({mode})"):
        with st.spinner("Processing…"):
            result_file = process_file(input_file, mode,
                                       mapping_df if mode == "Mapping" else None)

            st.success("✅ Output Generated!")
            st.download_button(
                "📥 Download Output",
                data=result_file,
                file_name="output_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown("---")
st.caption("Built for Rubick.ai | By Vishnu Sai")
