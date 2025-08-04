import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ──────────────────────────────────────────────────────────────
# Adjust these indexes if your mapping sheet layout is different
CLIENT_COL    = 0    # column A   – client name (e.g. Celio, Zivame)
INPUT_COL     = 1    # column B   – header exactly as in the user file
OUTPUT_COL    = 2    # column C   – header you want in the template
TYPE_ROW3_COL = 3    # column D   – value for row 3 in "Types"
TYPE_ROW4_COL = 4    # column E   – value for row 4 in "Types"
# ──────────────────────────────────────────────────────────────

TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH  = "Mapping - Automation.xlsx"

@st.cache_data
def load_mapping():
    return pd.read_excel(MAPPING_PATH, sheet_name=0)

def process_file(input_file, mode:str, mapping_df:pd.DataFrame|None=None):
    """
    Returns (BytesIO workbook, list_of_clients_used).

    * Mapping mode → honours every mapping row (→ duplicates, extra columns).
    * Auto-Mapping → fills Types rows 3-4 automatically.
    """
    src_df = pd.read_excel(input_file)
    columns_meta = []          # will drive both Values & Types sheets

    if mode == "Mapping" and mapping_df is not None:
        # 1-to-many mapping (duplicates inserted directly after the source)
        for col in src_df.columns:
            matches = mapping_df[mapping_df.iloc[:, INPUT_COL] == col]

            if matches.empty:      # keep unmapped column as-is
                columns_meta.append({
                    "src": col, "out": col,
                    "row3": "Not Found", "row4": "Not Found",
                    "client": None
                })
            else:                  # add one entry per matching mapping row
                for _, row in matches.iterrows():
                    columns_meta.append({
                        "src":  col,
                        "out":  (row.iloc[OUTPUT_COL] if pd.notna(row.iloc[OUTPUT_COL]) else col),
                        "row3": row.iloc[TYPE_ROW3_COL],
                        "row4": row.iloc[TYPE_ROW4_COL],
                        "client": row.iloc[CLIENT_COL]
                    })
    else:   # ─────────────── Auto-Mapping ───────────────
        for col in src_df.columns:
            # determine datatype for row 4
            dtype = "imageurlarray" if "image" in col.lower() else "string"
            columns_meta.append({
                "src": col, "out": col,
                "row3": "mandatory",
                "row4": dtype,
                "client": None         # not used in auto-mapping
            })

    # ─────────── Build the workbook from columns_meta ───────────
    wb        = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_values = wb["Values"]
    ws_types  = wb["Types"]

    # ► Values sheet
    for j, meta in enumerate(columns_meta, start=1):
        ws_values.cell(row=1, column=j, value=meta["out"])
        for i, val in enumerate(src_df[meta["src"]].tolist(), start=2):
            ws_values.cell(row=i, column=j, value=val)

    # ► Types sheet
    for j, meta in enumerate(columns_meta, start=2):
        ws_types.cell(row=1, column=j, value=meta["out"])
        ws_types.cell(row=2, column=j, value=meta["out"])
        ws_types.cell(row=3, column=j, value=meta["row3"])
        ws_types.cell(row=4, column=j, value=meta["row4"])

    # collect clients actually used (only meaningful for Mapping mode)
    clients_used = sorted({m["client"] for m in columns_meta if m["client"]})

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out, clients_used


# ───────────────────────────── Streamlit UI ─────────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("📊 SKU Template Automation Tool")

mode        = st.selectbox("Select Mode", ["Mapping", "Auto-Mapping"])
input_file  = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if input_file:
    if st.button(f"Generate Output ({mode})"):
        with st.spinner("Processing…"):
            mapping_df = load_mapping() if mode == "Mapping" else None
            result_file, clients = process_file(input_file, mode, mapping_df)

            st.success("✅ Output Generated!")
            st.download_button(
                "📥 Download Output",
                data=result_file,
                file_name="output_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            if mode == "Mapping" and clients:
                st.info(f"🔖 Mapping applied for: **{', '.join(clients)}**")

st.markdown("---")
st.caption("Built for Rubick.ai | By Vishnu Sai")
