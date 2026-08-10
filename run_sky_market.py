from zaevochnik_engine.pipeline import build_zaevochnik
from zaevochnik_engine.config import VALIDATION_PROMPTS, build_excel_styles

SOURCE_FILE = r"C:\Users\Пользователь\Desktop\Справочники\Справочник SKY MARKET.xlsm"
OUTPUT_FILE = r"C:\Users\Пользователь\Desktop\Заявочники АЗС\Заявочник SKY MARKET.xlsx"
SHEET_NAME = "Заявочник"
HEADER_START_ROW = 5

FILTER_COLUMN = None
EXCLUDE_VALUE = None

# Статусы, строки с которыми будут полностью удалены
EXCLUDE_STATUSES = ["Вывод"]

# Список временно создаваемых колонок для физического удаления в конце
COLUMNS_TO_DROP_FINALLY = ["Склад", "Группа товара", "Дата статуса"]

COLUMN_MAPPING = {
    "type_col": "Заказ",
    "qty_col": "Кол-во заказа",
    "mult_col": "В упаковке",
    "total_col": "Итого (шт)",
    "weight_col": "Вес"
}
TOTAL_WEIGHT_NAME = "Итого (вес)"

# Здесь остались ТОЛЬКО значения, которые отличают сеть PARTNER NEFT
# от остальных. Общая часть (row_heights, шрифты regular/bold, alignments,
# общие поля top_header, bg_new, border) лежит в config.COMMON_EXCEL_STYLES
# и подмешивается автоматически функцией build_excel_styles().
EXCEL_STYLES = build_excel_styles({
    "colors": {
        "primary": "F8309E",   # Основной корпоративный цвет (заливка шапки таблицы)
        "bg_pack": "FFE5FE"    # Фон для товаров, заказываемых УПАКОВКАМИ
    },
    "fonts": {
        "header": {"name": "Calibri", "size": 11, "bold": True, "color": "FEFFFF"}
    },
    "top_header": {
        "bg_color": "8D54EA",          # Фон всей верхней шапки (строки 1-4)
        "text_color_yellow": "FAE116", # Цвет для дат и подписей
        "text_color_white": "FEFFFF",  # Цвет для названия компании
        "company_name": "SKY MARKET",
        "fuel_station": "MARKET №"
    }
})

if __name__ == "__main__":
    print("--- Запуск генерации заявочника для сети Sky Market ---")
    build_zaevochnik(
        source_path=SOURCE_FILE,
        output_path=OUTPUT_FILE,
        sheet_name=SHEET_NAME,
        header_row=HEADER_START_ROW,
        filter_col=FILTER_COLUMN,
        filter_val=EXCLUDE_VALUE,
        drop_columns_finally=COLUMNS_TO_DROP_FINALLY,
        column_mapping=COLUMN_MAPPING,
        weight_column_name=TOTAL_WEIGHT_NAME,
        excel_styles=EXCEL_STYLES,
        validation_prompts=VALIDATION_PROMPTS,
        exclude_statuses=EXCLUDE_STATUSES
    )
