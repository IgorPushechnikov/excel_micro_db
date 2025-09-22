# src/storage/styles.py
"""
Модуль для работы со стилями в хранилище проекта Excel Micro DB.
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
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case).
    """
    if not font_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py (snake_case)
    # Убираем 'id' из списка, так как он генерируется автоматически
    columns = [
        'name', 'sz', 'b', 'i', 'u', 'strike', 'color_theme', 'color_type', 'color_rgb', 'color_tint',
        'vert_align', 'scheme', 'family', 'charset'
    ]
    placeholders = ', '.join(['?' for _ in columns])
    # ИСПРАВЛЕНО: Используем правильные имена столбцов для условия WHERE
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO fonts ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM fonts WHERE {select_conditions}"
    # ИСПРАВЛЕНО: Получаем значения по правильным ключам
    values = [font_attrs.get(key) for key in columns]
    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_fill(cursor, fill_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующей заливки или создает новую.
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case).
    """
    if not fill_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py (snake_case)
    columns = [
        'patternType', 'fg_color_theme', 'fg_color_type', 'fg_color_rgb', 'fg_color_tint',
        'bg_color_theme', 'bg_color_type', 'bg_color_rgb', 'bg_color_tint'
    ]
    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO pattern_fills ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM pattern_fills WHERE {select_conditions}"
    # ИСПРАВЛЕНО: Получаем значения по правильным ключам
    values = [fill_attrs.get(key) for key in columns]
    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_border(cursor, border_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующих границ или создает новые.
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case).
    """
    if not border_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py (snake_case)
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
    # ИСПРАВЛЕНО: Получаем значения по правильным ключам
    values = [border_attrs.get(key) for key in columns]
    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_alignment(cursor, align_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего выравнивания или создает новое.
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case).
    """
    if not align_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py (snake_case)
    columns = [
        'horizontal', 'vertical', 'text_rotation', 'wrap_text', 'shrink_to_fit', 'indent',
        'relative_indent', 'justify_last_line', 'reading_order'
    ]
    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO alignments ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM alignments WHERE {select_conditions}"
    # ИСПРАВЛЕНО: Получаем значения по правильным ключам
    values = [align_attrs.get(key) for key in columns]
    cursor.execute(insert_sql, values)
    cursor.execute(select_sql, values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_protection(cursor, prot_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующей защиты или создает новую.
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case).
    """
    if not prot_attrs:
        return None
    # Приводим BOOLEAN значения к INTEGER
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py (snake_case)
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
    cursor.execute(select_sql, values)
    row = cursor.fetchone()
    return row[0] if row else None

def _get_or_create_cell_style(cursor, style_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего стиля ячейки или создает новый.
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case).
    """
    if not style_attrs:
        return None
    # Извлекаем атрибуты для каждого компонента стиля
    # ИСПРАВЛЕНО: Используем правильные префиксы ключей (snake_case)
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
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py (snake_case)
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
    cursor.execute(select_sql, values)
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
            # ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case)
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
        # === ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (snake_case) ===
        cursor.execute('''
            SELECT
                sr.range_address,
                f.name, f.sz, f.b, f.i, f.u, f.strike, f.color_rgb, f.color_theme, f.color_tint, f.vert_align, f.scheme, f.family, f.charset,
                pf.patternType, pf.fg_color_rgb, pf.fg_color_theme, pf.fg_color_tint, pf.bg_color_rgb, pf.bg_color_theme, pf.bg_color_tint,
                b.left_style, b.left_color_rgb, b.right_style, b.right_color_rgb, b.top_style, b.top_color_rgb,
                b.bottom_style, b.bottom_color_rgb, b.diagonal_style, b.diagonal_color_rgb, b.diagonalUp, b.diagonalDown, b.outline,
                a.horizontal, a.vertical, a.text_rotation, a.wrap_text, a.shrink_to_fit, a.indent,
                a.relative_indent, a.justify_last_line, a.reading_order,
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
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов шрифта (snake_case)
            font_keys = ["name", "sz", "b", "i", "u", "strike", "color_rgb", "color_theme", "color_tint", "vert_align", "scheme", "family", "charset"]
            for i, key in enumerate(font_keys):
                if row[i+1] is not None: # +1 потому что range_address это 0
                    style_attrs[f"font_{key}"] = row[i+1]
            # Fill (индексы 15-21)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов заливки (snake_case)
            fill_keys = ["patternType", "fg_color_rgb", "fg_color_theme", "fg_color_tint", "bg_color_rgb", "bg_color_theme", "bg_color_tint"]
            for i, key in enumerate(fill_keys):
                if row[i+15] is not None:
                    style_attrs[f"fill_{key}"] = row[i+15]
            # Border (индексы 22-34)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов границ (snake_case)
            border_keys = ["left_style", "left_color_rgb", "right_style", "right_color_rgb", "top_style", "top_color_rgb",
                           "bottom_style", "bottom_color_rgb", "diagonal_style", "diagonal_color_rgb", "diagonalUp", "diagonalDown", "outline"]
            for i, key in enumerate(border_keys):
                if row[i+22] is not None:
                    style_attrs[f"border_{key}"] = row[i+22]
            # Alignment (индексы 35-44)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов выравнивания (snake_case)
            align_keys = ["horizontal", "vertical", "text_rotation", "wrap_text", "shrink_to_fit", "indent",
                          "relative_indent", "justify_last_line", "reading_order"]
            for i, key in enumerate(align_keys):
                if row[i+35] is not None:
                    style_attrs[f"alignment_{key}"] = row[i+35]
            # Protection (индексы 45-46)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов защиты (snake_case)
            prot_keys = ["locked", "hidden"]
            for i, key in enumerate(prot_keys):
                if row[i+45] is not None:
                    style_attrs[f"protection_{key}"] = row[i+45]
            # Cell Style main attrs (индексы 47-49)
            # ИСПРАВЛЕНО: Используем правильные ключи для основных атрибутов стиля (snake_case)
            cs_keys = ["num_fmt_id", "xf_id", "quote_prefix"]
            for i, key in enumerate(cs_keys):
                if row[i+47] is not None:
                    style_attrs[key] = row[i+47]
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
