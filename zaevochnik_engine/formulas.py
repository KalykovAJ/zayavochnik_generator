from openpyxl.utils import get_column_letter

def get_excel_col_letter(worksheet, start_row: int, column_name: str) -> str:
    """Определяет текущую букву столбца в Excel по его названию в шапке."""
    for col_idx in range(1, worksheet.max_column + 1):
        if str(worksheet.cell(row=start_row, column=col_idx).value).strip() == column_name:
            return get_column_letter(col_idx)
    raise ValueError(f"Столбец '{column_name}' не найден в итоговой структуре Excel.")

def apply_dynamic_formulas(worksheet, start_row: int, end_row: int, header_row: int, column_mapping: dict, weight_column_name: str):
    """Записывает формулы на основе пересчитанных букв колонок."""
    f_type = get_excel_col_letter(worksheet, header_row, column_mapping["type_col"])
    f_qty = get_excel_col_letter(worksheet, header_row, column_mapping["qty_col"])
    f_mult = get_excel_col_letter(worksheet, header_row, column_mapping["mult_col"])
    f_total = get_excel_col_letter(worksheet, header_row, column_mapping["total_col"])
    f_weight = get_excel_col_letter(worksheet, header_row, column_mapping["weight_col"])
    f_total_weight = get_excel_col_letter(worksheet, header_row, weight_column_name)

    for i in range(start_row, end_row + 1):
        formula_qty = (
            f'=IF(ISNUMBER(SEARCH("упаковка", {f_type}{i})), {f_qty}{i}*{f_mult}{i},'
            f'IF(ISNUMBER(SEARCH("штука", {f_type}{i})), {f_qty}{i},'
            f'IF(ISNUMBER(SEARCH("напрямую", {f_type}{i})), 0, "")))'
        )
        formula_weight = f'=IF({f_total}{i}<>"",{f_total}{i}*{f_weight}{i},"")'

        worksheet[f"{f_total}{i}"] = formula_qty
        worksheet[f"{f_total_weight}{i}"] = formula_weight