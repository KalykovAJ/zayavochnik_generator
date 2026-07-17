from zaevochnik_engine.pipeline import build_zaevochnik
from zaevochnik_engine.config import VALIDATION_PROMPTS, build_excel_styles

SOURCE_FILE = r"C:\Users\Пользователь\Desktop\Справочники\Справочник БП.xlsm"
OUTPUT_FILE = r"C:\Users\Пользователь\Desktop\Заявочники АЗС\Заявочник БП мини.xlsx"
#SOURCE_FILE = r"C:\Users\ajkal\OneDrive\Desktop\Справочники\Справочник БП.xlsm"
#OUTPUT_FILE = r"C:\Users\ajkal\OneDrive\Desktop\Заявочники АЗС\Заявочник БП.xlsx"
SHEET_NAME = "Заявочник"
HEADER_START_ROW = 5

FILTER_COLUMN = "Мини АЗС"
EXCLUDE_VALUE = "Да"

# Статусы, строки с которыми будут полностью удалены
EXCLUDE_STATUSES = ["Вывод"]

# ТРЕБОВАНИЕ 1: Исключили "Статус" из списка удаления, теперь колонка остается в финальном файле
COLUMNS_TO_DROP_FINALLY = ["Склад", "Группа товара", "Дата статуса", "Мини АЗС"]

COLUMN_MAPPING = {
    "type_col": "Заказ",
    "qty_col": "Кол-во заказа",
    "mult_col": "В упаковке",
    "total_col": "Итого (шт)",
    "weight_col": "Вес"
}
TOTAL_WEIGHT_NAME = "Итого (вес)"

# Здесь остались ТОЛЬКО значения, которые отличают сеть Bishkek Petroleum
# от остальных. Всё одинаковое (row_heights, шрифты regular/bold, alignments,
# общие поля top_header, bg_new, border) лежит в config.COMMON_EXCEL_STYLES
# и подмешивается автоматически функцией build_excel_styles().
EXCEL_STYLES = build_excel_styles({
    "colors": {
        "primary": "FAE116",
        "bg_pack": "DDF0E1"
    },
    "fonts": {
        "header": {"name": "Calibri", "size": 11, "bold": True, "color": "10AA19"}
    },
    "top_header": {
        "bg_color": "10AA19",
        "text_color_yellow": "FAE116",
        "text_color_white": "FEFFFD",
        "company_name": "Bishkek Petroleum Mini",
        "fuel_station": "Мини АЗС №"
    }
})

if __name__ == "__main__":
    print("--- Запуск генерации заявочника для сети АЗС ---")
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
