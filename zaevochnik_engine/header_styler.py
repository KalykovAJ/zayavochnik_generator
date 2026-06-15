import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from zaevochnik_engine.formulas import get_excel_col_letter


def apply_top_header_and_protection(worksheet, start_row: int, end_row: int, column_mapping: dict,
                                    weight_column_name: str, excel_styles: dict):
    """
    Формирует и стилизует верхнюю шапку (строки 1-4) на основе скриншота без внутренних границ,
    добавляет префикс формата веса, а также включает защиту ячеек для всего листа.
    """
    cfg = excel_styles["top_header"]
    bg_fill = PatternFill(start_color=cfg["bg_color"], end_color=cfg["bg_color"], fill_type="solid")

    # Стили шрифтов Cambria
    font_company = Font(name=cfg["font_name"],
                        size=cfg["company_name_font_size"] if "company_name_font_size" in cfg else cfg[
                            "company_font_size"], bold=True, color=cfg["text_color_white"])
    font_yellow = Font(name=cfg["font_name"], size=cfg["label_font_size"], bold=True, color=cfg["text_color_yellow"])
    font_white = Font(name=cfg["font_name"], size=cfg["label_font_size"], bold=True, color=cfg["text_color_white"])

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # 1. Снимаем старые границы и красим фон во всем блоке А1:I4
    for r in range(1, 5):
        for c in range(1, 10):
            cell = worksheet.cell(row=r, column=c)
            cell.fill = bg_fill
            cell.border = Border()  # Убираем границы
            cell.protection = Protection(locked=True)  # По умолчанию заблокировано

    # 2. Наполнение структуры текстом и объединение по макету скриншота
    # Строка 1: Последнее обновление
    worksheet.merge_cells("A1:B1")
    worksheet["A1"] = f"Последнее обновление заявочника: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}"
    worksheet["A1"].font = font_yellow
    worksheet["A1"].alignment = align_left

    # Центральный блок: Мунай Пром (строки 1-4, колонки C-E)
    worksheet.merge_cells("C1:E4")
    worksheet["C1"] = cfg["company_name"]
    worksheet["C1"].font = font_company
    worksheet["C1"].alignment = align_center

    # Строка 3: АЗС № и поле ввода
    worksheet["A3"] = "АЗС №"
    worksheet["A3"].font = font_white
    worksheet["A3"].alignment = align_right

    worksheet["B3"] = "здесь операторы пишут номер азс"
    worksheet["B3"].font = font_white
    worksheet["B3"].alignment = align_left
    worksheet["B3"].protection = Protection(locked=False)  # Разблокировано для ввода!

    # Строка 4: Дата заявки
    worksheet["A4"] = "Дата заявки:"
    worksheet["A4"].font = font_white
    worksheet["A4"].alignment = align_right

    worksheet["B4"] = datetime.datetime.now().strftime("%d.%m.%Y")
    worksheet["B4"].font = font_white
    worksheet["B4"].alignment = align_left

    # Правый блок: Общий вес и его сумма (колонки H и I)
    worksheet.merge_cells("H1:H4")
    worksheet["H1"] = "Общий вес\n(кг)"
    worksheet["H1"].font = font_white
    worksheet["H1"].alignment = align_center

    worksheet.merge_cells("I1:I4")
    worksheet["I1"].font = font_white
    worksheet["I1"].alignment = align_center

    # Записываем формулу суммы веса динамически на основе буквы колонки
    f_total_weight_letter = get_excel_col_letter(worksheet, start_row, weight_column_name)
    worksheet["I1"] = f"=SUM({f_total_weight_letter}{start_row + 1}:{f_total_weight_letter}{end_row})"

    # 3. Настройка числовых форматов: суффикс "кг" для всей колонки веса
    for row in range(start_row + 1, end_row + 1):
        cell = worksheet[f"{f_total_weight_letter}{row}"]
        cell.number_format = '#,##0.00" кг"'
    # Суффикс для ячейки общей суммы в шапке
    worksheet["I1"].number_format = '#,##0.00" кг"'

    # 4. Тотальная защита листа, кроме столбца "Кол-во заказа"
    qty_letter = get_excel_col_letter(worksheet, start_row, column_mapping["qty_col"])

    # Разблокируем ячейки данных в столбце "Кол-во заказа"
    for row in range(start_row + 1, end_row + 1):
        worksheet[f"{qty_letter}{row}"].protection = Protection(locked=False)

    # Включаем аппаратную защиту самого листа в Excel
    worksheet.protection.password = cfg["sheet_password"]
    worksheet.protection.enable()