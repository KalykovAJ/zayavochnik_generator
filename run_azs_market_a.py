from zaevochnik_engine.pipeline import build_zaevochnik

SOURCE_FILE = r"C:\Users\ajkal\OneDrive\Desktop\Справочник МП.xlsx"
OUTPUT_FILE = r"C:\Users\ajkal\OneDrive\Desktop\Результат_Заявочник.xlsx"
SHEET_NAME = "Заявочник"
HEADER_START_ROW = 5

FILTER_COLUMN = "Склад"
EXCLUDE_VALUE = ["Автохимия"]

# Статусы, строки с которыми будут полностью удалены
EXCLUDE_STATUSES = ["Вывод"]

# Список временно создаваемых колонок для физического удаления в конце
COLUMNS_TO_DROP_FINALLY = ["Склад", "Группа товара", "Статус", "Дата статуса"]

COLUMN_MAPPING = {
    "type_col": "Заказ",
    "qty_col": "Кол-во заказа",
    "mult_col": "В упаковке",
    "total_col": "Итого (шт)",
    "weight_col": "Вес"
}
TOTAL_WEIGHT_NAME = "Итоговый вес"

# Полная конфигурация стилей для конкретной сети АЗС
EXCEL_STYLES = {
    # ─── РАЗМЕРЫ СТРОК (В ПУНКТАХ) ───────────────────────────────────────
    "row_heights": {
        "header": 28,  # Высота строки для шапки таблицы (чтобы текст не зажимало)
        "data": 21  # Высота всех строк с товарами (комфортная для чтения)
    },

    # ─── ЦВЕТОВАЯ ПАЛИТРА (HEX-КОДЫ RRGGBB) ─────────────────────────────
    "colors": {
        "primary": "F9630E",  # Основной корпоративный цвет (заливка шапки таблицы)
        "bg_pack": "E6F0FA",  # Светло-голубой фон для товаров, заказываемых УПАКОВКАМИ
        "bg_direct": "F9EFEA",  # Светло-бежевый фон для товаров ПРЯМОЙ ПОСТАВКИ
        "bg_suspended": "FFC7CE",  # Светло-красный фон для товаров с приостановленной отгрузкой
        "bg_new": "FFEB9C",  # Светло-желтый фон для товаров со статусом "Новинка"
        "border": "D9D9D9"  # Нейтральный серый цвет для внутренних линий сетки
    },

    # ─── НАСТРОЙКИ ШРИФТОВ ДЛЯ ТАБЛИЦЫ ──────────────────────────────────
    "fonts": {
        "header": {"name": "Calibri", "size": 11, "bold": True, "color": "FCFAFB"},
        "regular": {"name": "Calibri", "size": 11, "bold": False},
        "bold": {"name": "Calibri", "size": 11, "bold": True}
    },

    # ─── ВЫРАВНИВАНИЕ ТЕКСТА В ЯЧЕЙКАХ ──────────────────────────────────
    "alignments": {
        "header": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "default": {"horizontal": "center", "vertical": "center"},
        "text_left": {"horizontal": "left", "vertical": "center"}
    },

    # ─── НОВЫЕ НАСТРОЙКИ ДЛЯ СТИЛИЗАЦИИ И ЗАЩИТЫ ВЕРХНЕЙ ШАПКИ ──────────
    "top_header": {
        "bg_color": "1D71B8",         # Синий фон всей верхней шапки (строки 1-4)
        "font_name": "Cambria",        # Шрифт по ТЗ
        "text_color_yellow": "FFD54E", # Цвет для дат и подписей
        "text_color_white": "FFFFFF",  # Цвет для названия компании
        "border": "D9D9D9",
        "company_name": "Мунай Пром",
        "company_font_size": 26,       # Размер шрифта для бренда
        "label_font_size": 12,         # Размер для остальных текстов шапки
        "sheet_password": "1526"       # Пароль защиты книги/листа
    }
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
        "message": "Отгрузка данного товара временно переноправлена. Ввод заблокирован."
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