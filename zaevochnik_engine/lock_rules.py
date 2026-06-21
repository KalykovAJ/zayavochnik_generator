"""
Модуль определения "замораживающих" правил для строк заявочника.

Объединяет логику:
  - п.1 ТЗ: статус товара (например "Приостановлено") -> строка полностью
    блокируется и закрашивается заданным цветом;
  - п.2 ТЗ: тип заказа (например "Напрямую поставщику") -> то же самое
    условие, что и в п.1.

Правила хранятся как списки словарей в config.py (LOCKED_STATUS_RULES,
LOCKED_ORDER_TYPE_RULES), поэтому добавление нового статуса или типа
заказа со своим цветом не требует изменения кода - достаточно дополнить
соответствующий список в config.py.
"""

from typing import Optional, Tuple


def _match_color(value_text: str, rules: list) -> Optional[str]:
    """
    Ищет первое правило из списка, чьё ключевое слово встречается
    в value_text (без учёта регистра). Возвращает цвет правила или None.
    """
    if not rules:
        return None

    text = (value_text or "").strip().lower()
    if not text:
        return None

    for rule in rules:
        keywords = rule.get("keywords") or []
        color = rule.get("color")
        if not keywords or not color:
            continue
        if any(str(keyword).strip().lower() in text for keyword in keywords):
            return color

    return None


def resolve_row_lock(status_text: str, type_text: str,
                     status_rules: list, order_type_rules: list) -> Tuple[bool, Optional[str]]:
    """
    Определяет, должна ли строка быть полностью заблокирована для
    редактирования и каким цветом её нужно закрасить.

    Приоритет проверки: сначала статус (п.1), затем тип заказа (п.2) -
    это сохраняет исходное поведение кода (is_suspended проверялся раньше
    is_direct).

    Возвращает (is_locked, fill_color). Если ни одно правило не сработало -
    (False, None).
    """
    color = _match_color(status_text, status_rules)
    if color:
        return True, color

    color = _match_color(type_text, order_type_rules)
    if color:
        return True, color

    return False, None


def is_locked_row(status_text: str, type_text: str,
                  status_rules: list, order_type_rules: list) -> bool:
    """Удобный шорткат для мест, где нужен только булевый признак блокировки."""
    locked, _ = resolve_row_lock(status_text, type_text, status_rules, order_type_rules)
    return locked
