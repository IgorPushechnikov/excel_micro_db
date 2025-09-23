# src/storage/styles.py
"""
Модуль для работы со стилиями в хранилище проекта Excel Micro DB.
"""

import sqlite3
import logging
from typing import List, Dict, Any, Optional

# from src.storage.schema import ... # Если потребуются какие-либо константы

logger = logging.getLogger(__name__)

# --- Вспомогательные функции (перенесены из database.py) ---

# Эти функции по-прежнему нужны внутри этого модуля

def _get_or_create_font(cursor, font_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего шрифта или создает новый.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not font_attrs:
        return None

    # ИСПРАВЛЕНО: Используем правильные имена столбцов из реальной БД
    # color_tint, color_rgb, color_theme, color_type добавлены
    # Стиль именования столбцов: camelCase
    columns = [
        'name', 'sz', 'b', 'i', 'u', 'strike', 'color_theme', 'color_type', 'color_rgb', 'color_tint',
        'vert_align', 'scheme', 'family', 'charset'
    ]

    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO fonts ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM fonts WHERE {select_conditions}"

    # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов шрифта
    # Ключи в словаре font_attrs должны соответствовать именам столбцов в БД или оригинальному имени атрибута openpyxl
    # Предполагается, что font_attrs передается с ключами в стиле snake_case или оригинальными именами openpyxl
    # Нужно адаптировать ключи к именам столбцов в БД (camelCase)
    db_keys_mapping = {
        'name': 'name',
        'sz': 'sz',
        'b': 'b',
        'i': 'i',
        'u': 'u',
        'strike': 'strike',
        'color_theme': 'color_theme',
        'color_type': 'color_type',
        'color_rgb': 'color_rgb',
        'color_tint': 'color_tint',
        'vert_align': 'vert_align',
        'scheme': 'scheme',
        'family': 'family',
        'charset': 'charset'
    }
    
    # Формируем список значений в порядке столбцов БД
    values = [font_attrs.get(py_key) for py_key in db_keys_mapping.keys()]

    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values + values) # Условия WHERE для IS NULL тоже проверяются
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_fill(cursor, fill_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующей заливки или создает новую.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not fill_attrs:
        return None

    # ИСПРАВЛЕНО: Используем правильные имена столбцов из реальной БД
    # fgColor_tint, fgColor_rgb, fgColor_theme, fgColor_type и т.д. добавлены
    # Стиль именования столбцов: camelCase с префиксом типа цвета
    columns = [
        'patternType', 'fgColor_theme', 'fgColor_type', 'fgColor_rgb', 'fgColor_tint',
        'bgColor_theme', 'bgColor_type', 'bgColor_rgb', 'bgColor_tint'
    ]

    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO pattern_fills ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM pattern_fills WHERE {select_conditions}"

    # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов заливки
    db_keys_mapping = {
        'patternType': 'patternType',
        'fgColor_theme': 'fgColor_theme',
        'fgColor_type': 'fgColor_type',
        'fgColor_rgb': 'fgColor_rgb',
        'fgColor_tint': 'fgColor_tint',
        'bgColor_theme': 'bgColor_theme',
        'bgColor_type': 'bgColor_type',
        'bgColor_rgb': 'bgColor_rgb',
        'bgColor_tint': 'bgColor_tint'
    }
    
    values = [fill_attrs.get(py_key) for py_key in db_keys_mapping.keys()]

    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values + values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_border(cursor, border_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующих границ или создает новые.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not border_attrs:
        return None

    # ИСПРАВЛЕНО: Используем правильные имена столбцов из реальной БД
    # color_tint, color_rgb, color_theme, color_type добавлены для всех сторон
    # Стиль именования столбцов: camelCase с префиксом стороны
    columns = [
        'left_style', 'left_color_theme', 'left_color_type', 'left_color_rgb', 'left_color_tint',
        'right_style', 'right_color_theme', 'right_color_type', 'right_color_rgb', 'right_color_tint',
        'top_style', 'top_color_theme', 'top_color_type', 'top_color_rgb', 'top_color_tint',
        'bottom_style', 'bottom_color_theme', 'bottom_color_type', 'bottom_color_rgb', 'bottom_color_tint',
        'diagonal_style', 'diagonal_color_theme', 'diagonal_color_type', 'diagonal_color_rgb', 'diagonal_color_tint',
        'diagonalUp', 'diagonalDown', 'outline'
    ]

    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO borders ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM borders WHERE {select_conditions}"

    # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов границ
    db_keys_mapping = {
        'left_style': 'left_style',
        'left_color_theme': 'left_color_theme',
        'left_color_type': 'left_color_type',
        'left_color_rgb': 'left_color_rgb',
        'left_color_tint': 'left_color_tint',
        'right_style': 'right_style',
        'right_color_theme': 'right_color_theme',
        'right_color_type': 'right_color_type',
        'right_color_rgb': 'right_color_rgb',
        'right_color_tint': 'right_color_tint',
        'top_style': 'top_style',
        'top_color_theme': 'top_color_theme',
        'top_color_type': 'top_color_type',
        'top_color_rgb': 'top_color_rgb',
        'top_color_tint': 'top_color_tint',
        'bottom_style': 'bottom_style',
        'bottom_color_theme': 'bottom_color_theme',
        'bottom_color_type': 'bottom_color_type',
        'bottom_color_rgb': 'bottom_color_rgb',
        'bottom_color_tint': 'bottom_color_tint',
        'diagonal_style': 'diagonal_style',
        'diagonal_color_theme': 'diagonal_color_theme',
        'diagonal_color_type': 'diagonal_color_type',
        'diagonal_color_rgb': 'diagonal_color_rgb',
        'diagonal_color_tint': 'diagonal_color_tint',
        'diagonalUp': 'diagonalUp',
        'diagonalDown': 'diagonalDown',
        'outline': 'outline'
    }
    
    values = [border_attrs.get(py_key) for py_key in db_keys_mapping.keys()]

    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values + values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_alignment(cursor, align_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего выравнивания или создает новое.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not align_attrs:
        return None

    # ИСПРАВЛЕНО: Используем правильные имена столбцов из реальной БД
    # wrapText, shrinkToFit, relativeIndent, justifyLastLine, readingOrder
    # Стиль именования столбцов: snake_case (кроме text_rotation, которое уже исправлено в схеме)
    columns = [
        'horizontal', 'vertical', 'text_rotation', 'wrapText', 'shrinkToFit', 'indent',
        'relativeIndent', 'justifyLastLine', 'readingOrder'
    ]

    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO alignments ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM alignments WHERE {select_conditions}"

    # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов выравнивания
    # Предполагается, что ключи в align_attrs соответствуют именам столбцов в БД
    db_keys_mapping = {
        'horizontal': 'horizontal',
        'vertical': 'vertical',
        'text_rotation': 'text_rotation', # Уже исправлено в схеме
        'wrapText': 'wrapText',
        'shrinkToFit': 'shrinkToFit',
        'indent': 'indent',
        'relativeIndent': 'relativeIndent',
        'justifyLastLine': 'justifyLastLine',
        'readingOrder': 'readingOrder'
    }
    
    values = [align_attrs.get(py_key) for py_key in db_keys_mapping.keys()]

    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values + values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_protection(cursor, prot_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующей защиты или создает новую.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not prot_attrs:
        return None

    # Приводим BOOLEAN значения к INTEGER
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из реальной БД
    # Стиль именования столбцов: snake_case
    prot_attrs_converted = {
        "locked": int(prot_attrs.get("locked", 1)) if prot_attrs.get("locked") is not None else 1,
        "hidden": int(prot_attrs.get("hidden", 0)) if prot_attrs.get("hidden") is not None else 0,
    }

    columns = list(prot_attrs_converted.keys())
    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO protections ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM protections WHERE {select_conditions}"

    values = [prot_attrs_converted.get(col) for col in columns]

    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values + values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_cell_style(cursor, style_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего стиля ячейки или создает новый.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not style_attrs:
        return None

    # Извлекаем атрибуты для каждого компонента стиля
    # ИСПРАВЛЕНО: Используем правильные префиксы ключей
    # Предполагается, что style_attrs содержит ключи с префиксами, например, 'font_name', 'fill_patternType'
    font_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('font_')}
    fill_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('fill_')}
    border_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('border_')}
    align_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('alignment_')}
    prot_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('protection_')}

    # Получаем или создаем ID для каждого компонента
    font_id = _get_or_create_font(cursor, font_attrs)
    fill_id = _get_or_create_fill(cursor, fill_attrs)
    border_id = _get_or_create_border(cursor, border_attrs)
    alignment_id = _get_or_create_alignment(cursor, align_attrs)
    protection_id = _get_or_create_protection(cursor, prot_attrs)

    # Атрибуты самого стиля (те, что не входят в подкомпоненты)
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из реальной БД
    # Стиль именования столбцов: snake_case
    style_main_attrs = {
        "font_id": font_id,
        "fill_id": fill_id,
        "border_id": border_id,
        "alignment_id": alignment_id,
        "protection_id": protection_id,
        "num_fmt_id": style_attrs.get("num_fmt_id"),
        "xf_id": style_attrs.get("xf_id"),
        "quote_prefix": int(style_attrs.get("quote_prefix", 0)) if style_attrs.get("quote_prefix") is not None else None,
    }

    columns = list(style_main_attrs.keys())
    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO cell_styles ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM cell_styles WHERE {select_conditions}"

    values = [style_main_attrs.get(col) for col in columns]

    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values + values)
    row = cursor.fetchone()
    return row[0] if row else None

# --- Основные функции работы со стилями ---

def save_sheet_styles(connection: sqlite3.Connection, sheet_id: int, styled_ranges_data: List[Dict[str, Any]]) -> bool:
    """
    Сохраняет уникальные стили и их применение к диапазонам на листе.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа в БД.
        styled_ranges_data (List[Dict[str, Any]]): Список словарей, где каждый словарь
            содержит 'style_attributes' (dict) и 'range_address' (str).

    Returns:
        bool: True, если данные сохранены успешно, False в случае ошибки.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения стилей.")
        return False

    try:
        cursor = connection.cursor()
        for style_range_info in styled_ranges_data:
            style_attrs = style_range_info.get("style_attributes", {})
            range_addr = style_range_info.get("range_address", "")

            if not style_attrs or not range_addr:
                logger.warning(f"Пропущен стиль/диапазон: {style_range_info}")
                continue

            # 1. Получаем или создаем ID стиля
            style_id = _get_or_create_cell_style(cursor, style_attrs)
            if style_id is None:
                logger.error(f"Не удалось получить или создать стиль для: {style_attrs}")
                continue

            # 2. Сохраняем связь стиль-диапазон
            # ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД
            cursor.execute('''
                INSERT OR IGNORE INTO styled_ranges (sheet_id, style_id, range_address)
                VALUES (?, ?, ?)
            ''', (sheet_id, style_id, range_addr))

        connection.commit()
        logger.info(f"Стили для листа ID {sheet_id} успешно сохранены.")
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении стилей для листа ID {sheet_id}: {e}")
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении стилей для листа ID {sheet_id}: {e}")
        connection.rollback()
        return False

def load_sheet_styles(connection: sqlite3.Connection, sheet_id: int) -> List[Dict[str, Any]]:
    """
    Загружает стили и диапазоны для указанного листа.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа в БД.

    Returns:
        List[Dict[str, Any]]: Список словарей с 'style_attributes' и 'range_address'.
    """
    styles_data = []
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки стилей.")
        return styles_data

    try:
        cursor = connection.cursor()

        # Запрашиваем связанные стили и диапазоны
        # Это сложный JOIN, который собирает всю информацию о стиле
        # === ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД ===
        # Используем camelCase для столбцов из fonts, pattern_fills, borders
        # Используем snake_case для столбцов из alignments, protections, cell_styles
        cursor.execute('''
            SELECT
                sr.range_address,
                f.name, f.sz, f.b, f.i, f.u, f.strike, f.color_rgb, f.color_theme, f.color_tint, f.vert_align, f.scheme, f.family, f.charset,
                pf.patternType, pf.fgColor_rgb, pf.fgColor_theme, pf.fgColor_tint, pf.bgColor_rgb, pf.bgColor_theme, pf.bgColor_tint,
                b.left_style, b.left_color_rgb, b.right_style, b.right_color_rgb, b.top_style, b.top_color_rgb,
                b.bottom_style, b.bottom_color_rgb, b.diagonal_style, b.diagonal_color_rgb, b.diagonalUp, b.diagonalDown, b.outline,
                a.horizontal, a.vertical, a.text_rotation, a.wrapText, a.shrinkToFit, a.indent,
                a.relativeIndent, a.justifyLastLine, a.readingOrder,
                p.locked, p.hidden,
                cs.num_fmt_id, cs.xf_id, cs.quote_prefix
            FROM styled_ranges sr
            LEFT JOIN cell_styles cs ON sr.style_id = cs.id
            LEFT JOIN fonts f ON cs.font_id = f.id
            LEFT JOIN pattern_fills pf ON cs.fill_id = pf.id
            LEFT JOIN borders b ON cs.border_id = b.id
            LEFT JOIN alignments a ON cs.alignment_id = a.id
            LEFT JOIN protections p ON cs.protection_id = p.id
            WHERE sr.sheet_id = ?
        ''', (sheet_id,))
        # =================================================

        rows = cursor.fetchall()
        for row in rows:
            range_addr = row[0]
            # Собираем атрибуты стиля из результата запроса

            style_attrs = {}

            # Font (индексы 1-14)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов шрифта
            # Префикс 'font_' для атрибутов шрифта
            font_keys = ["name", "sz", "b", "i", "u", "strike", "color_rgb", "color_theme", "color_tint", "vert_align", "scheme", "family", "charset"]
            for i, key in enumerate(font_keys):
                if row[i + 1] is not None:  # +1 потому что range_address это 0
                    style_attrs[f"font_{key}"] = row[i + 1]

            # Fill (индексы 15-21)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов заливки
            # Префикс 'fill_' для атрибутов заливки
            fill_keys = ["patternType", "fgColor_rgb", "fgColor_theme", "fgColor_tint", "bgColor_rgb", "bgColor_theme", "bgColor_tint"]
            for i, key in enumerate(fill_keys):
                if row[i + 15] is not None:
                    style_attrs[f"fill_{key}"] = row[i + 15]

            # Border (индексы 22-34)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов границ
            # Префикс 'border_' для атрибутов границ
            # ВАЖНО: Нужно убедиться, что порядок ключей соответствует порядку в SELECT
            border_keys = [
                "left_style", "left_color_rgb", # 22, 23
                "right_style", "right_color_rgb", # 24, 25
                "top_style", "top_color_rgb", # 26, 27
                "bottom_style", "bottom_color_rgb", # 28, 29
                "diagonal_style", "diagonal_color_rgb", # 30, 31
                "diagonalUp", "diagonalDown", "outline" # 32, 33, 34
            ]
            for i, key in enumerate(border_keys):
                if row[i + 22] is not None:
                    style_attrs[f"border_{key}"] = row[i + 22]

            # Alignment (индексы 35-44)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов выравнивания
            # Префикс 'alignment_' для атрибутов выравнивания
            align_keys = ["horizontal", "vertical", "text_rotation", "wrapText", "shrinkToFit", "indent",
                          "relativeIndent", "justifyLastLine", "readingOrder"]
            for i, key in enumerate(align_keys):
                if row[i + 35] is not None:
                    style_attrs[f"alignment_{key}"] = row[i + 35]

            # Protection (индексы 45-46)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов защиты
            # Префикс 'protection_' для атрибутов защиты
            prot_keys = ["locked", "hidden"]
            for i, key in enumerate(prot_keys):
                if row[i + 45] is not None:
                    style_attrs[f"protection_{key}"] = row[i + 45]

            # Cell Style main attrs (индексы 47-49)
            # ИСПРАВЛЕНО: Используем правильные ключи для основных атрибутов стиля
            # Без префикса, так как это атрибуты самого стиля
            cs_keys = ["num_fmt_id", "xf_id", "quote_prefix"]
            for i, key in enumerate(cs_keys):
                if row[i + 47] is not None:
                    style_attrs[key] = row[i + 47]

            styles_data.append({
                "style_attributes": style_attrs,
                "range_address": range_addr
            })

        logger.info(f"Загружено {len(styles_data)} стилей для листа ID {sheet_id}.")
        return styles_data

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке стилей для листа ID {sheet_id}: {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке стилей для листа ID {sheet_id}: {e}")
        return []