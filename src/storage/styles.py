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

def _get_or_create_font(cursor, font_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего шрифта или создает новый.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    ИСПРАВЛЕНО: Количество передаваемых значений в SQL-запросах.
    """
    if not font_attrs:
        return None

    columns = [
        'name', 'sz', 'b', 'i', 'u', 'strike', 'color_theme', 'color_type', 'color_rgb', 'color_tint',
        'vert_align', 'scheme', 'family', 'charset'
    ]

    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO fonts ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM fonts WHERE {select_conditions}"

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
    
    values = [font_attrs.get(key) for key in db_keys_mapping.keys()]
    # ИСПРАВЛЕНО: values_for_where должен содержать значения только один раз
    values_for_where = values

    try:
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values_for_where) # ИСПРАВЛЕНО
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite в _get_or_create_font: {e}")
        # Можно добавить raise или обработку по желанию
        return None

def _get_or_create_fill(cursor, fill_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующей заливки или создает новую.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    ИСПРАВЛЕНО: Количество передаваемых значений в SQL-запросах.
    """
    if not fill_attrs:
        return None

    columns = [
        'patternType', 'fgColor_theme', 'fgColor_type', 'fgColor_rgb', 'fgColor_tint',
        'bgColor_theme', 'bgColor_type', 'bgColor_rgb', 'bgColor_tint'
    ]

    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO pattern_fills ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM pattern_fills WHERE {select_conditions}"

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
    
    values = [fill_attrs.get(key) for key in db_keys_mapping.keys()]
    # ИСПРАВЛЕНО: values_for_where должен содержать значения только один раз
    values_for_where = values

    try:
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values_for_where) # ИСПРАВЛЕНО
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite в _get_or_create_fill: {e}")
        return None

def _get_or_create_border(cursor, border_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующих границ или создает новые.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    ИСПРАВЛЕНО: Количество передаваемых значений в SQL-запросах.
    """
    if not border_attrs:
        return None

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
    
    values = [border_attrs.get(key) for key in db_keys_mapping.keys()]
    # ИСПРАВЛЕНО: values_for_where должен содержать значения только один раз
    values_for_where = values

    try:
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values_for_where) # ИСПРАВЛЕНО
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite в _get_or_create_border: {e}")
        return None

def _get_or_create_alignment(cursor, align_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего выравнивания или создает новое.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not align_attrs:
        return None

    columns = [
        'horizontal', 'vertical', 'text_rotation', 'wrapText', 'shrinkToFit', 'indent',
        'relativeIndent', 'justifyLastLine', 'readingOrder'
    ]

    placeholders = ', '.join(['?' for _ in columns])
    select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
    insert_sql = f"INSERT OR IGNORE INTO alignments ({', '.join(columns)}) VALUES ({placeholders})"
    select_sql = f"SELECT id FROM alignments WHERE {select_conditions}"

    db_keys_mapping = {
        'horizontal': 'horizontal',
        'vertical': 'vertical',
        'text_rotation': 'text_rotation',
        'wrapText': 'wrapText',
        'shrinkToFit': 'shrinkToFit',
        'indent': 'indent',
        'relativeIndent': 'relativeIndent',
        'justifyLastLine': 'justifyLastLine',
        'readingOrder': 'readingOrder'
    }
    
    values = [align_attrs.get(key) for key in db_keys_mapping.keys()]
    # Для условия WHERE нужны только значения один раз
    values_for_where = values

    try:
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values_for_where)
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite в _get_or_create_alignment: {e}")
        return None

def _get_or_create_protection(cursor, prot_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующей защиты или создает новую.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not prot_attrs:
        return None

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
    # Для условия WHERE нужны только значения один раз
    values_for_where = values

    try:
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values_for_where)
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite в _get_or_create_protection: {e}")
        return None

def _get_or_create_cell_style(cursor, style_attrs: Dict[str, Any]) -> Optional[int]:
    """
    Получает ID существующего стиля ячейки или создает новый.
    ИСПРАВЛЕНО: Имена столбцов соответствуют реальной схеме БД.
    """
    if not style_attrs:
        return None

    font_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('font_')}
    fill_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('fill_')}
    border_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('border_')}
    align_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('alignment_')}
    prot_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('protection_')}

    font_id = _get_or_create_font(cursor, font_attrs)
    fill_id = _get_or_create_fill(cursor, fill_attrs)
    border_id = _get_or_create_border(cursor, border_attrs)
    alignment_id = _get_or_create_alignment(cursor, align_attrs)
    protection_id = _get_or_create_protection(cursor, prot_attrs)

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
    # Для условия WHERE нужны только значения один раз
    values_for_where = values

    try:
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values_for_where)
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite в _get_or_create_cell_style: {e}")
        return None

# --- Основные функции работы со стилями ---

def save_sheet_styles(connection: sqlite3.Connection, sheet_id: int, styled_ranges_data: List[Dict[str, Any]]) -> bool:
    """
    Сохраняет уникальные стили и их применение к диапазонам на листе.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения стилей.")
        return False

    try:
        cursor = connection.cursor()
        success_count = 0
        for style_range_info in styled_ranges_data:
            style_attrs = style_range_info.get("style_attributes", {})
            range_addr = style_range_info.get("range_address", "")

            if not style_attrs or not range_addr:
                logger.warning(f"Пропущен стиль/диапазон: {style_range_info}")
                continue

            style_id = _get_or_create_cell_style(cursor, style_attrs)
            if style_id is None:
                logger.error(f"Не удалось получить или создать стиль для: {style_attrs}")
                # Можно продолжить или прервать, зависит от требований
                continue 

            cursor.execute('''
                INSERT OR IGNORE INTO styled_ranges (sheet_id, style_id, range_address)
                VALUES (?, ?, ?)
            ''', (sheet_id, style_id, range_addr))
            success_count += 1

        connection.commit()
        logger.info(f"Стили для листа ID {sheet_id} успешно сохранены. Обработано записей: {success_count}/{len(styled_ranges_data)}.")
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
    """
    styles_data = []
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки стилей.")
        return styles_data

    try:
        cursor = connection.cursor()

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

        rows = cursor.fetchall()
        for row in rows:
            range_addr = row[0]
            style_attrs = {}

            # Font (индексы 1-14)
            font_keys = ["name", "sz", "b", "i", "u", "strike", "color_rgb", "color_theme", "color_tint", "vert_align", "scheme", "family", "charset"]
            for i, key in enumerate(font_keys):
                if row[i + 1] is not None:
                    style_attrs[f"font_{key}"] = row[i + 1]

            # Fill (индексы 15-21)
            fill_keys = ["patternType", "fgColor_rgb", "fgColor_theme", "fgColor_tint", "bgColor_rgb", "bgColor_theme", "bgColor_tint"]
            for i, key in enumerate(fill_keys):
                if row[i + 15] is not None:
                    style_attrs[f"fill_{key}"] = row[i + 15]

            # Border (индексы 22-34)
            border_keys = [
                "left_style", "left_color_rgb",
                "right_style", "right_color_rgb",
                "top_style", "top_color_rgb",
                "bottom_style", "bottom_color_rgb",
                "diagonal_style", "diagonal_color_rgb",
                "diagonalUp", "diagonalDown", "outline"
            ]
            for i, key in enumerate(border_keys):
                if row[i + 22] is not None:
                    style_attrs[f"border_{key}"] = row[i + 22]

            # Alignment (индексы 35-44)
            align_keys = ["horizontal", "vertical", "text_rotation", "wrapText", "shrinkToFit", "indent",
                          "relativeIndent", "justifyLastLine", "readingOrder"]
            for i, key in enumerate(align_keys):
                if row[i + 35] is not None:
                    style_attrs[f"alignment_{key}"] = row[i + 35]

            # Protection (индексы 45-46)
            prot_keys = ["locked", "hidden"]
            for i, key in enumerate(prot_keys):
                if row[i + 45] is not None:
                    style_attrs[f"protection_{key}"] = row[i + 45]

            # Cell Style main attrs (индексы 47-49)
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
