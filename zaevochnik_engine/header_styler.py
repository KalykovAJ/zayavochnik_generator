import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from zaevochnik_engine.formulas import get_excel_col_letter


def apply_top_header_and_protection(worksheet, start_row: int, end_row: int, column_mapping: dict,
                                    weight_column_name: str, excel_styles: dict):
    """
    Формирует и стилизует верхнюю шапку (строки 1-4).
    Правый информационный блок позиционируется динамически на основе финального max_column.
    """
    cfg = excel_styles["top_header"]
    bg_fill = PatternFill(start_color=cfg["bg_color"], end_color=cfg["bg_color"], fill_type="solid")

    font_company = Font(name=cfg["font_name"],
                        size=cfg["company_name_font_size"] if "company_name_font_size" in cfg else cfg[
                            "company_font_size"], bold=True, color=cfg["text_color_white"])
    font_yellow = Font(name=cfg["font_name"], size=cfg["label_font_size"], bold=True, color=cfg["text_color_yellow"])
    font_white = Font(name=cfg["font_name"], size=cfg["label_font_size"], bold=True, color=cfg["text_color_white"])

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    border_color = cfg.get("border", "000000")
    thin_side = Side(border_style="thin", color=border_color)

    max_col = worksheet.max_column

    # Требование 1: Границы верхней брендированной плашки строятся динамически до актуального конца таблицы
    for r in range(1, 5):
        for c in range(1, max_col + 1):
            cell = worksheet.cell(row=r, column=c)
            cell.fill = bg_fill
            cell.protection = Protection(locked=True)

            cell.border = Border(
                top=thin_side if r == 1 else None,
                bottom=thin_side if r == 4 else None,
                left=thin_side if c == 1 else None,
                right=thin_side if c == max_col else None
            )

    # Наполнение левой и центральной части шапки
    worksheet.merge_cells("A1:B1")
    worksheet["A1"] = f"Последнее обновление заявочника: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}"
    worksheet["A1"].font = font_yellow
    worksheet["A1"].alignment = align_left

    worksheet.merge_cells("C1:F4")
    worksheet["C1"] = cfg["company_name"]
    worksheet["C1"].font = font_company
    worksheet["C1"].alignment = align_center

    worksheet["A2"] = "АЗС №"
    worksheet["A2"].font = font_white
    worksheet["A2"].alignment = align_right

    worksheet["B2"] = "[ введите номер ]"
    worksheet["B2"].font = font_white
    worksheet["B2"].alignment = align_left
    worksheet["B2"].protection = Protection(locked=False)

    worksheet["A3"] = "Дата заявки:"
    worksheet["A3"].font = font_white
    worksheet["A3"].alignment = align_right

    worksheet["B3"] = "=TODAY()"
    worksheet["B3"].font = font_white
    worksheet["B3"].alignment = align_left
    worksheet["B3"].number_format = "DD.MM.YYYY"  

    # Требование 1 & Универсальность: Динамический расчет координат правого блока веса
    col_lbl_letter = get_column_letter(max_col - 1)
    col_val_letter = get_column_letter(max_col)


    # Ячейка подписи названия
    worksheet.merge_cells(f"{col_lbl_letter}1:{col_lbl_letter}4")
    worksheet[f"{col_lbl_letter}1"] = "Общий вес:"
    worksheet[f"{col_lbl_letter}1"].font = font_white
    worksheet[f"{col_lbl_letter}1"].alignment = align_center

    # Ячейка формулы суммирования
    worksheet.merge_cells(f"{col_val_letter}1:{col_val_letter}4")
    f_total_weight_letter = get_excel_col_letter(worksheet, start_row, weight_column_name)
    worksheet[f"{col_val_letter}1"] = f"=SUM({f_total_weight_letter}{start_row + 1}:{f_total_weight_letter}{end_row})"
    worksheet[f"{col_val_letter}1"].font = font_white
    worksheet[f"{col_val_letter}1"].alignment = align_center
    worksheet[f"{col_val_letter}1"].number_format = '#,##0.00" кг"'

    # Считывание по актуальным буквам колонок
    qty_letter = get_excel_col_letter(worksheet, start_row, column_mapping["qty_col"])
    type_letter = get_excel_col_letter(worksheet, start_row, column_mapping["type_col"])
    status_letter = get_excel_col_letter(worksheet, start_row, "Статус")

    for row in range(start_row + 1, end_row + 1):
        status_val = str(worksheet[f"{status_letter}{row}"].value or "").strip().lower()
        type_val = str(worksheet[f"{type_letter}{row}"].value or "").strip().lower()

        is_suspended = ("приостановл" in status_val)
        is_direct = ("прямая" in type_val or "напрямую" in type_val)

        if is_suspended or is_direct:
            # Требование 3: Гарантированно блокируем ВСЕ ячейки строки
            for col_idx in range(1, max_col + 1):
                worksheet.cell(row=row, column=col_idx).protection = Protection(locked=True)
        else:
            # Штатный режим
            worksheet[f"{qty_letter}{row}"].protection = Protection(locked=False)
            for col_idx in range(1, max_col + 1):
                if get_column_letter(col_idx) != qty_letter:
                    worksheet.cell(row=row, column=col_idx).protection = Protection(locked=True)

    # Автоподбор ширины колонок шапки (A и B)
    for col_idx in [1, 2]:
        col_letter = get_column_letter(col_idx)
        max_header_len = 0
        for row_idx in range(1, 5):
            cell_val = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_val is not None and not str(cell_val).startswith("="):
                max_header_len = max(max_header_len, len(str(cell_val)))

        header_calculated_width = min(max_header_len + 5, 20)
        current_width = worksheet.column_dimensions[col_letter].width or 10
        worksheet.column_dimensions[col_letter].width = max(current_width, header_calculated_width)

    # Динамический автоподбор высоты строк шапки
    for r in range(2, 5):
        max_row_height = 22
        for c in range(1, max_col + 1):
            cell = worksheet.cell(row=r, column=c)
            if cell.value is not None and not str(cell.value).startswith("=") and c <= 2:
                val_str = str(cell.value)
                col_letter = get_column_letter(c)
                col_width = worksheet.column_dimensions[col_letter].width or 12
                if len(val_str) > col_width:
                    lines_count = (len(val_str) // int(col_width)) + 1
                    calculated_height = lines_count * 16
                    max_row_height = max(max_row_height, calculated_height)
        worksheet.row_dimensions[r].height = max_row_height

    worksheet.row_dimensions[1].height = 24

    # Аппаратная защита книги
    worksheet.protection.password = cfg["sheet_password"]
    worksheet.protection.selectLockedCells = True
    worksheet.protection.selectUnlockedCells = False
    worksheet.protection.enable()