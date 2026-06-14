from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


def apply_excel_styles(worksheet, start_row: int, end_row: int, column_mapping: dict, colors: dict,
                       validation_prompts: dict, row_statuses: dict):
    """Применяет стили оформления и настраивает контекстные подсказки на итоговую структуру."""

    primary_color = colors["primary"]
    styles = {
        "header_fill": PatternFill(start_color=primary_color, end_color=primary_color, fill_type="solid"),
        "header_font": Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
        "pack_fill": PatternFill(start_color=colors["bg_pack"], fill_type="solid"),
        "direct_fill": PatternFill(start_color=colors["bg_direct"], fill_type="solid"),
        "status_suspended_fill": PatternFill(start_color="FFC7CE", fill_type="solid"),
        "status_new_fill": PatternFill(start_color="FFEB9C", fill_type="solid"),
        "default_fill": PatternFill(fill_type=None),
        "regular_font": Font(name="Calibri", size=11, bold=False),
        "bold_font": Font(name="Calibri", size=11, bold=True),
        "align_center": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "align_left": Alignment(horizontal="left", vertical="center"),
        "border_data": Border(
            left=Side(border_style="thin", color=colors["border"]),
            right=Side(border_style="thin", color=colors["border"]),
            top=Side(border_style="thin", color=colors["border"]),
            bottom=Side(border_style="thin", color=colors["border"])
        ),
        "border_header": Border(
            left=Side(border_style="thin", color=colors["border"]),
            right=Side(border_style="thin", color=colors["border"]),
            top=Side(border_style="thin", color=colors["border"]),
            bottom=Side(border_style="medium", color=primary_color)
        )
    }

    # Ищем колонки в уже очищенной таблице (без "Статуса")
    headers = {str(worksheet.cell(row=start_row, column=i).value).strip(): i for i in
               range(1, worksheet.max_column + 1)}
    type_idx = headers.get(column_mapping["type_col"])
    qty_idx = headers.get(column_mapping["qty_col"])

    # Создаем 4 изолированных правила валидации, чтобы настройки сообщений не перемешивались
    dv_unit = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0")
    dv_pack = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0")
    dv_direct = DataValidation(type="whole", operator="equal", formula1="0")
    dv_suspended = DataValidation(type="whole", operator="equal", formula1="0")

    # Инициализируем тексты подсказок из конфигуратора (Пункты 1, 2, 3, 4)
    for key, dv in [("unit", dv_unit), ("pack", dv_pack), ("direct", dv_direct), ("suspended", dv_suspended)]:
        p = validation_prompts.get(key, {"title": "Внимание", "message": "Заполните поле"})
        dv.promptTitle = p["title"]
        dv.prompt = p["message"]
        dv.errorTitle = "Ошибка ввода"
        dv.error = "Введенное значение не соответствует правилам."
        dv.showInputMessage = True
        dv.showErrorMessage = True
        worksheet.add_data_validation(dv)

    # Красим шапку
    worksheet.row_dimensions[start_row].height = 28
    for col_idx in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=start_row, column=col_idx)
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]
        cell.border = styles["border_header"]
        cell.alignment = styles["align_center"]

    # Итерируемся по строкам
    for row_idx in range(start_row + 1, end_row + 1):
        worksheet.row_dimensions[row_idx].height = 21

        is_pack, is_direct, is_unit = False, False, False
        is_suspended, is_new = False, False

        # ЧИТАЕМ СТАТУС ИЗ СЛОВАРЯ (Пункт 4) — колонка удалена, но данные у нас есть!
        status_text = row_statuses.get(row_idx, "")
        if "приостановл" in status_text:
            is_suspended = True
        elif "новинк" in status_text:
            is_new = True

        # Читаем тип заказа (Пункты 1, 2, 3)
        if type_idx:
            cell_type = worksheet.cell(row=row_idx, column=type_idx)
            text = str(cell_type.value or "").lower()
            is_direct = "напрямую" in text
            is_pack = "упаковк" in text
            is_unit = "штук" in text

        # Вычисляем цвет строки
        if is_suspended:
            current_fill = styles["status_suspended_fill"]
        elif is_new:
            current_fill = styles["status_new_fill"]
        elif is_direct:
            current_fill = styles["direct_fill"]
        elif is_pack or is_unit:
            current_fill = styles["pack_fill"]
        else:
            current_fill = styles["default_fill"]

        # Применяем подсказки и правила Data Validation к ячейке количества
        if qty_idx:
            cell_qty = worksheet.cell(row=row_idx, column=qty_idx)

            if is_suspended:
                dv_suspended.add(cell_qty)
                cell_qty.value = 0
            elif is_direct:
                dv_direct.add(cell_qty)
                cell_qty.value = 0
            elif is_pack:
                dv_pack.add(cell_qty)
                cell_qty.value = None
            elif is_unit:
                dv_unit.add(cell_qty)
                cell_qty.value = None

        # Оформление ячеек
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            header_name = str(worksheet.cell(row=start_row, column=col_idx).value or "").strip()
            header_val_lower = header_name.lower()

            if current_fill.fill_type:
                cell.fill = current_fill

            cell.font = styles["bold_font"] if "итого" in header_val_lower else styles["regular_font"]
            cell.border = styles["border_data"]

            if header_name == "Наименование":
                cell.alignment = styles["align_left"]
            else:
                cell.alignment = styles["align_center"]

    # Автоподбор ширины
    for col in worksheet.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col if not str(cell.value or "").startswith("="))
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)