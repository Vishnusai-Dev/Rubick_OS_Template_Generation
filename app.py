    # ────────── BUILD THE WORKBOOK ──────────
    wb        = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals   = wb["Values"]
    ws_types  = wb["Types"]

    for j, m in enumerate(columns_meta, start=1):
        header_display = clean_header(m["out"])

        # ── Values sheet (starts at col 1) ──
        ws_vals.cell(row=1, column=j, value=header_display)

        # copy values; cast to text for string / imageurlarray columns
        for i, v in enumerate(src_df[m["src"]].tolist(), start=2):
            cell = ws_vals.cell(row=i, column=j)

            if pd.isna(v):
                cell.value = None
                continue

            if str(m["row4"]).lower() in ("string", "imageurlarray"):
                cell.value = str(v)
                cell.number_format = "@"
            else:
                cell.value = v

        # ── Types sheet (starts at col 3 … keep first two template columns) ──
        tcol = j + 2                    # shift everything two columns right
        ws_types.cell(row=1, column=tcol, value=header_display)
        ws_types.cell(row=2, column=tcol, value=header_display)
        ws_types.cell(row=3, column=tcol, value=m["row3"])
        ws_types.cell(row=4, column=tcol, value=m["row4"])
