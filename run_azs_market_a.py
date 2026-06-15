from zaevochnik_engine.pipeline import build_zaevochnik

SOURCE_FILE = r"C:\Users\Пользователь\Desktop\Справочник МП.xlsm"
OUTPUT_FILE = r"C:\Users\Пользователь\Desktop\Результат_Заявочник.xlsx"
SHEET_NAME = "Заявочник"
HEADER_START_ROW = 5

FILTER_COLUMN = "Склад"
EXCLUDE_VALUE = ["Автохимия"]

# Статусы, строки с которыми будут полностью удалены (Пункт 2)
# Можно передать строку, список ["Вывод", "Другой Статус"] или None
EXCLUDE_STATUSES = ["Вывод"]

# Теперь статус физически удаляется, и это ничего не ломает! (Пункт 5)
COLUMNS_TO_DROP_FINALLY = ["Склад", "Группа товара", "Статус", "Дата статуса"]

COLUMN_MAPPING = {
    "type_col": "Заказ",
    "qty_col": "Кол-во заказа",
    "mult_col": "В упаковке",
    "total_col": "Итого (шт)",
    "weight_col": "Вес"
}
TOTAL_WEIGHT_NAME = "Итоговый вес (кг)"

COLORS = {
    "primary": "1D71B8",
    "bg_pack": "E6F0FA",
    "bg_direct": "F9EFEA",
    "border": "D9D9D9"
}

VALIDATION_PROMPTS = {
    "unit": {
        "title": "Заказ в штуках! 🔎",
        "message": "Внимание: заказ принимается строго в ШТУКАХ."
    },
    "pack": {
        "title": "Заказ в упаковках! 📦",
        "message": "Внимание: заказ принимается строго в УПАКОВКАХ."
    },
    "direct": {
        "title": "Прямая поставка! ➡️",
        "message": "Заказ оформляется напрямую у поставщика. Ввод заблокирован."
    },
    "suspended": {
        "title": "Отгрузка приостановлена! ⛔",
        "message": "Отгрузка данного товара временно приостановлена. Ввод заблокирован."
    }
}

if __name__ == "__main__":
    print("--- Запуск генерации для сети АЗС 'Альфа' ---")
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
        color_palette=COLORS,
        validation_prompts=VALIDATION_PROMPTS,
        exclude_statuses=EXCLUDE_STATUSES  # Передаем конфигурацию удаления строк
    )