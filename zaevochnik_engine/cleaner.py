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


def remove_rows_by_status(worksheet, start_row: int, total_rows: int, status_values: list) -> int:
    """
    Находит колонку 'Статус' и физически удаляет строки,
    если значение в них совпадает со списком нежелательных статусов.
    Возвращает обновленное количество строк (total_rows).
    """
    if not status_values:
        return total_rows

    headers = {str(worksheet.cell(row=start_row, column=i).value).strip(): i
               for i in range(1, worksheet.max_column + 1)}

    status_idx = headers.get("Статус")
    if not status_idx:
        return total_rows

    # Приводим искомые статусы к нижнему регистру для надежности
    exclude_statuses = [str(val).strip().lower() for val in status_values]

    # Идем снизу вверх, чтобы индексы строк не смещались при удалении
    for row_idx in range(total_rows, start_row, -1):
        cell_val = str(worksheet.cell(row=row_idx, column=status_idx).value or "").strip().lower()
        if cell_val in exclude_statuses:
            worksheet.delete_rows(row_idx, 1)
            total_rows -= 1

    return total_rows