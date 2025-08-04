def process_mapping_mode(input_file, mapping_df):
    input_df = pd.read_excel(input_file)
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_values = wb["Values"]
    ws_type = wb["Types"]

    # Process mapping to create expanded columns
    final_columns = []
    col_data_map = {}

    for col in input_df.columns:
        matches = mapping_df[mapping_df.iloc[:, 1] == col]
        if not matches.empty:
            for _, row in matches.iterrows():
                out_col_name = row[3]  # Column D (Row 3)
                final_columns.append(out_col_name)
                col_data_map[out_col_name] = input_df[col]
        else:
            final_columns.append(col)
            col_data_map[col] = input_df[col]

    # Create a new DataFrame for final output
    final_df = pd.DataFrame(col_data_map)

    # Write to "Values" tab
    for j, col_name in enumerate(final_df.columns, start=1):
        ws_values.cell(row=1, column=j, value=col_name)
    for i, row in enumerate(final_df.itertuples(index=False), start=2):
        for j, value in enumerate(row, start=1):
            ws_values.cell(row=i, column=j, value=value)

    # Write to "Types" tab (Row 1 to 4)
    for col_idx, col_name in enumerate(final_df.columns, start=2):
        ws_type.cell(row=1, column=col_idx, value=col_name)
        ws_type.cell(row=2, column=col_idx, value=col_name)

        match_row = mapping_df[mapping_df.iloc[:, 3] == col_name]
        if not match_row.empty:
            row3_val = match_row.iloc[0, 3]
            row4_val = match_row.iloc[0, 4]
            ws_type.cell(row=3, column=col_idx, value=row3_val)
            ws_type.cell(row=4, column=col_idx, value=row4_val)
        else:
            ws_type.cell(row=3, column=col_idx, value="Not Found")
            ws_type.cell(row=4, column=col_idx, value="Not Found")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
