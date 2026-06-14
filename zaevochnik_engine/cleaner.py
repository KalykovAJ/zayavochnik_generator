def remove_temporary_columns(worksheet, start_row: int, drop_columns: list):
    """Полностью физически удаляет ненужные колонки, двигаясь с конца таблицы к началу."""
    if not drop_columns:
        return

    headers = [str(worksheet.cell(row=start_row, column=i).value).strip()
               for i in range(1, worksheet.max_column + 1)]

    for col_idx in range(len(headers), 0, -1):
        header_name = headers[col_idx - 1]
        if header_name in drop_columns:
            worksheet.delete_cols(col_idx, 1)