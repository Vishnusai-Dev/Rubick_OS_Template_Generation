def process_mapping_mode(input_file, mapping_df):
    input_df = pd.read_excel(input_file)
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_values = wb["Values"]
    ws_type = wb["Types"]

    final_columns = []
    final_data = {}

    # For each row in the mapping file
    for _, row in mapping_df.iterrows():
        input_col = row[1]      # Column B: Input Header
        output_col = row[3]     # Column D: Output Header
        output_type = row[4]    # Column E: Data Type

        if input_col in input_df.columns:
            # If output column not already added, add it
            if output_col not in final_columns:
                final_columns.append(output_col)
                final_data[output_col] = input_df[input_col]  # Copy data

    # Append unmapped columns (i.e., columns in input but not in mapping) at the end
    for col in input_df.columns:
        if col not in final_columns:
            final_columns.append(col)
            final_data[col] = input_df[col]

    # Create final DataFrame
    final_df = pd.DataFrame(final_data)[final_columns]

    # --- Write to Values Sheet ---
    for j, col_name in enumerate(final_df.columns, start=1):
        ws_values.cell(row=1, column=j, value=col_name)
    for i, row in enumerate(final_df.itertuples(index=False), start=2):
        for j, value in enumerate(row, start=1):
            ws_values.cell(row=i, column=j, value=value)

    # --- Write to Types Sheet ---
    for col_idx, col_name in enumerate(final_df.columns, start=2):
        ws_type.cell(row=1, column=col_idx, value=col_name)
        ws_type.cell(row=2, column=col_idx, value=col_name)

        match_row = mapping_df[mapping_df.iloc[:, 3] == col_name]  # Output Header
        if not match_row.empty:
            ws_type.cell(row=3, column=col_idx, value=col_name)
            ws_type.cell(row=4, column=col_idx, value=match_row.iloc[0, 4])  # Type
        else:
            ws_type.cell(row=3, column=col_idx, value="Not Found")
            ws_type.cell(row=4, column=col_idx, value="Not Found")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
