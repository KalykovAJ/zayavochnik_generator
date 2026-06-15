from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


def apply_excel_styles(worksheet, start_row: int, end_row: int, column_mapping: dict,
                       excel_styles: dict, validation_prompts: dict, row_statuses: dict):
    """Применяет стили оформления, динамически собирая их из внешнего конфигуратора сетки."""

    colors = excel_styles["colors"]
    fonts_cfg = excel_styles["fonts"]
    align_cfg = excel_styles["alignments"]
    heights = excel_styles["row_heights"]

    # Динамическая сборка стилей openpyxl
    styles = {
        "header_fill": PatternFill(start_color=colors["primary"], end_color=colors["primary"], fill_type="solid"),
        "header_font": Font(**fonts_cfg["header"]),

        "pack_fill": PatternFill(start_color=colors["bg_pack"], fill_type="solid"),
        "direct_fill": PatternFill(start_color=colors["bg_direct"], fill_type="solid"),
        "status_suspended_fill": PatternFill(start_color=colors["bg_suspended"], fill_type="solid"),
        "status_new_fill": PatternFill(start_color=colors["bg_new"], fill_type="solid"),
        "default_fill": PatternFill(fill_type=None),

        "regular_font": Font(**fonts_cfg["regular"]),
        "bold_font": Font(**fonts_cfg["bold"]),

        "align_center": Alignment(**align_cfg["default"]),
        "align_left": Alignment(**align_cfg["text_left"]),
        "align_header": Alignment(**align_cfg["header"]),

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
            bottom=Side(border_style="medium", color=colors["primary"])
        )
    }

    # Ищем колонки в уже очищенной таблице
    headers = {str(worksheet.cell(row=start_row, column=i).value).strip(): i for i in
               range(1, worksheet.max_column + 1)}
    type_idx = headers.get(column_mapping["type_col"])
    qty_idx = headers.get(column_mapping["qty_col"])

    # Создаем правила Data Validation
    dv_unit = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0")
    dv_pack = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0")
    dv_direct = DataValidation(type="whole", operator="equal", formula1="0")
    dv_suspended = DataValidation(type="whole", operator="equal", formula1="0")

    for key, dv in [("unit", dv_unit), ("pack", dv_pack), ("direct", dv_direct), ("suspended", dv_suspended)]:
        p = validation_prompts.get(key, {"title": "Внимание", "message": "Заполните поле"})
        dv.promptTitle = p["title"]
        dv.prompt = p["message"]
        dv.errorTitle = "Ошибка ввода"
        dv.error = "Введенное значение не соответствует правилам."
        dv.showInputMessage = True
        dv.showErrorMessage = True
        worksheet.add_data_validation(dv)

    # Оформляем шапку
    worksheet.row_dimensions[start_row].height = heights["header"]
    for col_idx in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=start_row, column=col_idx)
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]
        cell.border = styles["border_header"]
        cell.alignment = styles["align_header"]

    # Итерируемся по строкам данных
    for row_idx in range(start_row + 1, end_row + 1):
        worksheet.row_dimensions[row_idx].height = heights["data"]

        is_pack, is_direct, is_unit = False, False, False
        is_suspended, is_new = False, False

        # Читаем статус из сохраненного слепка
        status_text = row_statuses.get(row_idx, "")
        if "приостановл" in status_text:
            is_suspended = True
        elif "новинк" in status_text:
            is_new = True

        # Читаем тип заказа и добавляем визуальную стрелку ➔
        if type_idx:
            cell_type = worksheet.cell(row=row_idx, column=type_idx)
            current_type_val = str(cell_type.value or "").strip()
            text = current_type_val.lower()

            is_direct = "напрямую" in text
            is_pack = "упаковк" in text
            is_unit = "штук" in text

            # Добавляем стрелку к тексту, если её там ещё нет
            if (is_pack or is_unit) and "➔" not in current_type_val:
                cell_type.value = f"{current_type_val} ➔"

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

        # Применяем подсказки и валидацию
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

        # Оформление ячеек строки
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

    # ─── АВТОПОДБОР ШИРИНЫ КОЛОНОК ПО ДАННЫМ ТАБЛИЦЫ ────────────────────
    for col_idx in range(1, worksheet.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0

        # Сканируем строго от start_row (заголовки таблицы) и ниже
        for row_idx in range(start_row, end_row + 1):
            cell_val = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_val is not None and not str(cell_val).startswith("="):
                max_len = max(max_len, len(str(cell_val)))

        worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)