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
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (CamelCase).
    """
    if not font_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py
    columns = [
        'name', 'sz', 'b', 'i', 'u', 'strike', 'colorTheme', 'colorType', 'colorRgb',
        'vertAlign', 'scheme', 'family', 'charset'
    ]
    # ИСПРАВЛЕНО: Убираем color_tint, так как его нет в fonts schema
    # columns = [
    #     'name', 'sz', 'b', 'i', 'u', 'strike', 'colorTheme', 'colorType', 'colorRgb', 'colorTint',
    #     'vertAlign', 'scheme', 'family', 'charset'
    # ]
    placeholders = ', '.join(['?' for _ in columns])
    # ИСПРАВЛЕНО: Используем правильные имена столбцов
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
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (CamelCase).
    """
    if not fill_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py
    columns = [
        'patternType', 'fgColorTheme', 'fgColorType', 'fgColorRgb', 'fgColorTint',
        'bgColorTheme', 'bgColorType', 'bgColorRgb', 'bgColorTint'
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
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (CamelCase).
    """
    if not border_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py
    columns = [
        'leftStyle', 'leftColorTheme', 'leftColorType', 'leftColorRgb', 'leftColorTint',
        'rightStyle', 'rightColorTheme', 'rightColorType', 'rightColorRgb', 'rightColorTint',
        'topStyle', 'topColorTheme', 'topColorType', 'topColorRgb', 'topColorTint',
        'bottomStyle', 'bottomColorTheme', 'bottomColorType', 'bottomColorRgb', 'bottomColorTint',
        'diagonalStyle', 'diagonalColorTheme', 'diagonalColorType', 'diagonalColorRgb', 'diagonalColorTint',
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
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (CamelCase).
    """
    if not align_attrs:
        return None
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py
    columns = [
        'horizontal', 'vertical', 'textRotation', 'wrapText', 'shrinkToFit', 'indent',
        'relativeIndent', 'justifyLastLine', 'readingOrder'
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
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (CamelCase).
    """
    if not prot_attrs:
        return None
    # Приводим BOOLEAN значения к INTEGER
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py
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
    ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py (CamelCase).
    """
    if not style_attrs:
        return None
    # Извлекаем атрибуты для каждого компонента стиля
    # ИСПРАВЛЕНО: Используем правильные префиксы ключей
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
    # ИСПРАВЛЕНО: Используем правильные имена столбцов из schema.py
    style_main_attrs = {
        "fontId": font_id,
        "fillId": fill_id,
        "borderId": border_id,
        "alignmentId": alignment_id,
        "protectionId": protection_id,
        "numFmtId": style_attrs.get("num_fmt_id"),
        "xfId": style_attrs.get("xf_id"),
        "quotePrefix": int(style_attrs.get("quote_prefix", 0)) if style_attrs.get("quote_prefix") is not None else None,
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
            # ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py
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
        # === ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py ===
        cursor.execute('''
            SELECT
                sr.range_address,
                f.name, f.sz, f.b, f.i, f.u, f.strike, f.colorRgb, f.colorTheme, f.colorTint, f.vertAlign, f.scheme,
                pf.patternType, pf.fgColorRgb, pf.fgColorTheme, pf.fgColorTint, pf.bgColorRgb, pf.bgColorTheme, pf.bgColorTint,
                b.leftStyle, b.leftColorRgb, b.rightStyle, b.rightColorRgb, b.topStyle, b.topColorRgb,
                b.bottomStyle, b.bottomColorRgb, b.diagonalStyle, b.diagonalColorRgb, b.diagonalUp, b.diagonalDown, b.outline,
                a.horizontal, a.vertical, a.textRotation, a.wrapText, a.shrinkToFit, a.indent,
                a.relativeIndent, a.justifyLastLine, a.readingOrder,
                p.locked, p.hidden,
                cs.numFmtId, cs.xfId, cs.quotePrefix
            FROM styled_ranges sr
            LEFT JOIN cell_styles cs ON sr.style_id = cs.id
            LEFT JOIN fonts f ON cs.fontId = f.id
            LEFT JOIN pattern_fills pf ON cs.fillId = pf.id
            LEFT JOIN borders b ON cs.borderId = b.id
            LEFT JOIN alignments a ON cs.alignmentId = a.id
            LEFT JOIN protections p ON cs.protectionId = p.id
            WHERE sr.sheet_id = ?
        ''', (sheet_id,))
        # =================================================
        rows = cursor.fetchall()
        for row in rows:
            range_addr = row[0]
            # Собираем атрибуты стиля из результата запроса
            style_attrs = {}
            # Font (индексы 1-11)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов шрифта
            font_keys = ["name", "sz", "b", "i", "u", "strike", "colorRgb", "colorTheme", "colorTint", "vertAlign", "scheme"]
            for i, key in enumerate(font_keys):
                if row[i+1] is not None: # +1 потому что range_address это 0
                    style_attrs[f"font_{key}"] = row[i+1]
            # Fill (индексы 12-18)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов заливки
            fill_keys = ["patternType", "fgColorRgb", "fgColorTheme", "fgColorTint", "bgColorRgb", "bgColorTheme", "bgColorTint"]
            for i, key in enumerate(fill_keys):
                if row[i+12] is not None:
                    style_attrs[f"fill_{key}"] = row[i+12]
            # Border (индексы 19-30)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов границ
            border_keys = ["leftStyle", "leftColorRgb", "rightStyle", "rightColorRgb", "topStyle", "topColorRgb",
                           "bottomStyle", "bottomColorRgb", "diagonalStyle", "diagonalColorRgb", "diagonalUp", "diagonalDown", "outline"]
            for i, key in enumerate(border_keys):
                if row[i+19] is not None:
                    style_attrs[f"border_{key}"] = row[i+19]
            # Alignment (индексы 31-41)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов выравнивания
            align_keys = ["horizontal", "vertical", "textRotation", "wrapText", "shrinkToFit", "indent",
                          "relativeIndent", "justifyLastLine", "readingOrder"]
            for i, key in enumerate(align_keys):
                if row[i+31] is not None:
                    style_attrs[f"alignment_{key}"] = row[i+31]
            # Protection (индексы 42-43)
            # ИСПРАВЛЕНО: Используем правильные ключи для атрибутов защиты
            prot_keys = ["locked", "hidden"]
            for i, key in enumerate(prot_keys):
                if row[i+42] is not None:
                    style_attrs[f"protection_{key}"] = row[i+42]
            # Cell Style main attrs (индексы 44-46)
            # ИСПРАВЛЕНО: Используем правильные ключи для основных атрибутов стиля
            cs_keys = ["numFmtId", "xfId", "quotePrefix"]
            for i, key in enumerate(cs_keys):
                if row[i+44] is not None:
                    style_attrs[key] = row[i+44]
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
