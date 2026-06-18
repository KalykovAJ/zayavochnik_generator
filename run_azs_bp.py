from zaevochnik_engine.pipeline import build_zaevochnik

SOURCE_FILE = r"C:\Users\Пользователь\Desktop\Справочники\Справочник БП.xlsm"
OUTPUT_FILE = r"C:\Users\Пользователь\Desktop\Заявочники АЗС\Заявочник БП.xlsx"
SHEET_NAME = "Заявочник"
HEADER_START_ROW = 5

FILTER_COLUMN = "Склад"
EXCLUDE_VALUE = [None]

# Статусы, строки с которыми будут полностью удалены
EXCLUDE_STATUSES = ["Вывод"]

# ТРЕБОВАНИЕ 1: Исключили "Статус" из списка удаления, теперь колонка остается в финальном файле
COLUMNS_TO_DROP_FINALLY = ["Склад", "Группа товара", "Дата статуса"]

COLUMN_MAPPING = {
    "type_col": "Заказ",
    "qty_col": "Кол-во заказа",
    "mult_col": "В упаковке",
    "total_col": "Итого (шт)",
    "weight_col": "Вес"
}
TOTAL_WEIGHT_NAME = "Вес (кг)"

# Полная конфигурация стилей для конкретной сети АЗС
EXCEL_STYLES = {
    "row_heights": {
        "header": 28,
        "data": 21
    },
    "colors": {
        "primary": "FAE116",
        "bg_pack": "DDF0E1",
        "bg_direct": "F9EFEA",
        "bg_suspended": "F44336",
        "bg_new": "FFEB9C",
        "border": "D9D9D9"
    },
    "fonts": {
        "header": {"name": "Calibri", "size": 11, "bold": True, "color": "10AA19"},
        "regular": {"name": "Calibri", "size": 11, "bold": False},
        "bold": {"name": "Calibri", "size": 11, "bold": True}
    },
    "alignments": {
        "header": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "default": {"horizontal": "center", "vertical": "center"},
        "text_left": {"horizontal": "left", "vertical": "center"}
    },
    "top_header": {
        "bg_color": "10AA19",
        "font_name": "Cambria",
        "text_color_yellow": "FAE116",
        "text_color_white": "FEFFFD",
        "border": "D9D9D9",
        "company_name": "Bishkek Petroleum",
        "company_font_size": 26,
        "label_font_size": 12,
        "sheet_password": "1526"
    }
}

# ТРЕБОВАНИЕ 4: Удалили правила DataValidation для "direct" и "suspended", так как они теперь жестко блокируются интерфейсом Excel
VALIDATION_PROMPTS = {
    "unit": {
        "title": "Заказ в штуках! 🔎",
        "message": "Внимание: заказ принимается строго в ШТУКАХ."
    },
    "pack": {
        "title": "Заказ в упаковках! 📦",
        "message": "Внимание: заказ принимается строго в УПАКОВКАХ."
    }
}

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