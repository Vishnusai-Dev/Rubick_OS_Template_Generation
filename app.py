def process_mapping_mode(input_file, mapping_df):
    input_df = pd.read_excel(input_file)

    # Start fresh DataFrame with original input
    output_df = input_df.copy()

    # Dictionary to hold added (Output Header → (Level, DataType))
    new_columns = {}

    for _, row in mapping_df.iterrows():
        input_col = row[1]  # Input Header
        output_col = row[2]  # Output Header
        level = row[3]
        dtype = row[4]

        if input_col in input_df.columns:
            # Add output_col ONLY if not already present
            if output_col not in output_df.columns:
                output_df[output_col] = input_df[input_col]
                new_columns[output_col] = (level, dtype)
            else:
                # Forcefully duplicate again with _1, _2 if name exists
                suffix = 1
                new_name = f"{output_col}_{suffix}"
                while new_name in output_df.columns:
                    suffix += 1
                    new_name = f"{output_col}_{suffix}"
                output_df[new_name] = input_df[input_col]
                new_columns[new_name] = (level, dtype)

    # Prepare Excel workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_values = wb.create_sheet("Values")
    ws_type = wb.create_sheet("Types")

    # Fill Values Sheet
    for j, col in enumerate(output_df.columns, start=1):
        ws_values.cell(row=1, column=j, value=col)
    for i, row in enumerate(output_df.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            ws_values.cell(row=i, column=j, value=val)

    # Fill Types Sheet
    for col_idx, col in enumerate(output_df.columns, start=2):
        ws_type.cell(row=1, column=col_idx, value=col)
        ws_type.cell(row=2, column=col_idx, value=col)
        if col in new_columns:
            ws_type.cell(row=3, column=col_idx, value=new_columns[col][0])  # Level
            ws_type.cell(row=4, column=col_idx, value=new_columns[col][1])  # DataType
        else:
            ws_type.cell(row=3, column=col_idx, value="Not Mapped")
            ws_type.cell(row=4, column=col_idx, value="string")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
