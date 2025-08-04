import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ──────────────────────────────────────────────────────────────
TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH  = "Mapping - Automation.xlsx"

# column names in the Mapping sheet
ATTR_COL   = "Attributes"             # header exactly as in user file
TARGET_COL = "Field Name"             # header you want in template
MAND_COL   = "Mandatory OR Not"       # → Types row-3
TYPE_COL   = "Field Type"             # → Types row-4
DUP_COL    = "Duplicates to be created"
# ──────────────────────────────────────────────────────────────


# ╭───────────────────────── helpers ─────────────────────────╮
@st.cache_data
def load_mapping():
    """Return (mapping_df, list_of_client_names)."""
    mapping_df  = pd.read_excel(MAPPING_PATH, sheet_name="Mapping")
    clients_raw = pd.read_excel(MAPPING_PATH,
                                sheet_name="Mapped Client Name",
                                header=None)          # treat every cell as data
    client_list = [str(x) for x in clients_raw.values.flatten() if pd.notna(x)]
    return mapping_df, client_list


def process_file(input_file, mode: str, mapping_df: pd.DataFrame | None = None):
    """Build workbook for either Mapping or Auto-Mapping mode."""
    src_df = pd.read_excel(input_file)

    # ——- assemble one dict per *output* column ————————————
    columns_meta = []
    if mode == "Mapping" and mapping_df is not None:
        for col in src_df.columns:
            matches = mapping_df[mapping_df[ATTR_COL] == col]

            # ► original column (always kept)
            if not matches.empty:
                ref_row = matches.iloc[0]
                row3, row4 = ref_row[MAND_COL], ref_row[TYPE_COL]
            else:
                row3, row4 = "Not Found", "Not Found"

            columns_meta.append({"src": col, "out": col,
                                 "row3": row3, "row4": row4})

            # ► duplicates flagged “Yes”
            for _, row in matches.iterrows():
                if str(row[DUP_COL]).strip().lower().startswith("yes"):
                    new_header = row[TARGET_COL] if pd.notna(row[TARGET_COL]) else col
                    if new_header == col:          # avoid 1:1 self-duplicates
                        continue
                    columns_meta.append({"src": col, "out": new_header,
                                         "row3": row[MAND_COL],
                                         "row4": row[TYPE_COL]})

    else:   # ——- Auto-Mapping ——————————————
        for col in src_df.columns:
            dtype = "imageurlarray" if "image" in col.lower() else "string"
            columns_meta.append({"src": col, "out": col,
                                 "row3": "mandatory", "row4": dtype})

    # ——- build workbook ——————————————
    wb         = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_values  = wb["Values"]
    ws_types   = wb["Types"]

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

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
# ╰─────────────────────────────────────────────────────────────╯


# ───────────────────────── Streamlit UI ─────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("📊 SKU Template Automation Tool")

# load mapping (cached)
mapping_df, client_names = load_mapping()

# show mapped-client list up front
st.info(f"🗂️  **Mapped clients available:** {', '.join(client_names)}")

mode       = st.selectbox("Select Mode", ["Mapping", "Auto-Mapping"])
input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if input_file:
    if st.button(f"Generate Output ({mode})"):
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
