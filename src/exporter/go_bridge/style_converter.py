# -*- coding: utf-8 -*-
"""
Модуль для преобразования структуры стиля из формата, сохранённого в БД (openpyxl),
в формат, пригодный для передачи в Go-экспортер через JSON и Pydantic-модель Style.

Цель: рекурсивно обойти словарь стиля и заменить недопустимые значения
(например, строки, описывающие объекты openpyxl) на примитивные типы данных.
"""
import logging
from typing import Any, Dict, Union, List, Tuple

logger = logging.getLogger(__name__)

# Специальные значения, которые могут быть в JSON-строках из openpyxl
OPENPYXL_OBJECT_PLACEHOLDER = "<openpyxl"
COLOR_OBJECT_PLACEHOLDER = "Parameters:\\nrgb="


def _clean_value(value: Any) -> Any:
    """
    Рекурсивно очищает значение в структуре стиля.

    Args:
        value: Значение для очистки (может быть словарь, список, строка, примитив).

    Returns:
        Очищенное значение (того же типа, но без недопустимых строк объектов).
    """
    if isinstance(value, dict):
        # Рекурсивно обрабатываем словарь
        cleaned_dict = {}
        for k, v in value.items():
            cleaned_v = _clean_value(v)
            cleaned_dict[k] = cleaned_v
        return cleaned_dict
    elif isinstance(value, list):
        # Рекурсивно обрабатываем список
        return [_clean_value(item) for item in value]
    elif isinstance(value, str):
        # Проверяем, является ли строка описанием объекта openpyxl
        if OPENPYXL_OBJECT_PLACEHOLDER in value and COLOR_OBJECT_PLACEHOLDER in value:
            # Пытаемся извлечь RGB из строки, если это Color object
            # Пример строки: "<openpyxl.styles.colors.Color object>\nParameters:\nrgb=FF0000, indexed=None, auto=None, theme=1, tint=0.0, type='theme'"
            # Цель: получить "FF0000" или "000000" по умолчанию
            rgb_start = value.find("rgb=")
            if rgb_start != -1:
                rgb_part = value[rgb_start + 4:]  # Пропускаем "rgb="
                # Ищем конец значения (до запятой или перевода строки)
                end_pos = rgb_part.find(",")
                if end_pos == -1:
                    end_pos = rgb_part.find("\n")
                if end_pos != -1:
                    rgb_code = rgb_part[:end_pos].strip().strip("'\"")
                    if rgb_code and len(rgb_code) == 6:
                        logger.debug(f"Извлечен цвет RGB: {rgb_code} из строки: {value[:50]}...")
                        return rgb_code
                    else:
                        logger.warning(f"Не удалось извлечь валидный RGB из строки: {value[:100]}...")

            # Если не удалось извлечь RGB или его нет, возвращаем "000000" или None
            logger.warning(f"Найден объект openpyxl в строке стиля, заменён на '000000': {value[:100]}...")
            return "000000"
        # Другие возможные объекты openpyxl можно обрабатывать здесь по аналогии
        # Пока обрабатываем только Color
        return value
    else:
        # Примитивные типы (int, float, bool, None) возвращаем как есть
        return value


def convert_openpyxl_style_to_go_format(style_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    Преобразует словарь стиля из формата openpyxl (хранящегося в БД) в формат,
    пригодный для передачи в Go-экспортер.

    Args:
        style_dict: Словарь стиля, полученный из БД (json.loads от style_attributes).

    Returns:
        Очищенный словарь стиля, готовый к использованию в Pydantic и Go.
        Возвращает пустой словарь в случае ошибки.
    """
    if not isinstance(style_dict, dict):
        logger.error(f"Ожидался словарь стиля, получен {type(style_dict)}: {style_dict}")
        return {}

    try:
        cleaned_style = _clean_value(style_dict)
        logger.debug(f"Стиль успешно преобразован: {cleaned_style}")
        return cleaned_style
    except Exception as e:
        logger.error(f"Ошибка при преобразовании стиля: {e}", exc_info=True)
        # В случае ошибки возвращаем пустой словарь, чтобы не ломать процесс
        return {}


# --- Пример использования (можно закомментировать или удалить после тестирования) ---
# if __name__ == "__main__":
#     import json
#     # Пример строки из БД
#     example_style_str = '{"font": {"color": {"rgb": "<openpyxl.styles.colors.Color object>\\nParameters:\\nrgb=FF0000, indexed=None, auto=None, theme=1, tint=0.0, type=\'theme\'"}}}'
#     example_style_dict = json.loads(example_style_str)
#     print("Оригинальный словарь:", example_style_dict)
#     cleaned = convert_openpyxl_style_to_go_format(example_style_dict)
#     print("Очищенный словарь:", cleaned)
#     print("JSON для Go:", json.dumps(cleaned, indent=2))
