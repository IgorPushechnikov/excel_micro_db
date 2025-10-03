# backend/constructor/widgets/simple_gui/db_data_fetcher.py
"""
Модуль для извлечения данных из БД проекта в формате, удобном для SimpleSheetModel.
"""
import sqlite3
import json
import re
from typing import Dict, Any, List, Tuple, Optional
import logging

from backend.utils.logger import get_logger

logger = get_logger(__name__)


def _column_letter_to_index(letter: str) -> int:
    """Преобразует букву столбца Excel (например, 'A', 'Z', 'AA') в 0-базовый индекс."""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1  # 0-based index


def _parse_cell_address(addr: str) -> Optional[Tuple[int, int]]:
    """Парсит адрес ячейки (например, A1) в индексы (row, col)."""
    try:
        col_part = ''.join(filter(str.isalpha, addr)).upper()
        row_part = ''.join(filter(str.isdigit, addr))

        if not row_part or not col_part:
            return None

        row_idx = int(row_part) - 1  # 1-based to 0-based
        col_idx = _column_letter_to_index(col_part)

        return row_idx, col_idx
    except:
        return None


def _parse_range_address(range_addr: str) -> Optional[Tuple[int, int, int, int]]:
    """Парсит адрес диапазона (например, A1:B2) в индексы (start_row, start_col, end_row, end_col)."""
    # Регулярное выражение для адреса одной ячейки или диапазона
    # Примеры: 'A1', 'Z10', 'AA1', 'A1:B2', 'ZZ100:AAA200'
    # Группы: 1-столбец_начала, 2-строка_начала, 3-столбец_конца, 4-строка_конца (опционально)
    range_pattern = re.compile(r'^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$')
    match = range_pattern.match(range_addr)
    if not match:
        logger.warning(f"_parse_range_address: Не удалось распознать формат адреса '{range_addr}'.")
        return None

    start_col_letter, start_row_str, end_col_letter, end_row_str = match.groups()

    # Если это одиночная ячейка, end_... будет None
    if end_col_letter is None or end_row_str is None:
        end_col_letter = start_col_letter
        end_row_str = start_row_str

    try:
        start_col_index = _column_letter_to_index(start_col_letter)
        start_row_index = int(start_row_str) - 1  # Excel 1-based -> Python 0-based
        end_col_index = _column_letter_to_index(end_col_letter)
        end_row_index = int(end_row_str) - 1      # Excel 1-based -> Python 0-based

        if start_col_index < 0 or start_row_index < 0 or end_col_index < 0 or end_row_index < 0:
            logger.warning(f"_parse_range_address: Некорректный индекс после парсинга {range_addr}: [{start_row_index}, {start_col_index}] - [{end_row_index}, {end_col_index}]")
            return None

        return start_row_index, start_col_index, end_row_index, end_col_index

    except (ValueError, TypeError) as e:
        logger.warning(f"_parse_range_address: Ошибка преобразования индексов для '{range_addr}': {e}")
        return None


def fetch_sheet_data(sheet_name: str, db_path: str) -> Tuple[List[List[Any]], Dict[Tuple[int, int], Dict[str, Any]]]:
    """
    Извлекает сырые данные и стили для листа из БД.

    Args:
        sheet_name (str): Имя листа.
        db_path (str): Путь к файлу БД.

    Returns:
        Tuple[List[List[Any]], Dict[Tuple[int, int], Dict[str, Any]]]:
        - 2D список значений.
        - Словарь стилей { (row, col): {"font_color": "#FF0000", ...} }.
    """
    rows_2d: List[List[Any]] = []
    styles_map: Dict[Tuple[int, int], Dict[str, Any]] = {}

    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # 1. Получаем ID листа
        cursor.execute("SELECT sheet_id FROM sheets WHERE name = ?", (sheet_name,))
        sheet_row = cursor.fetchone()
        if not sheet_row:
            logger.warning(f"fetch_sheet_data: Лист '{sheet_name}' не найден в БД {db_path}")
            return rows_2d, styles_map
        sheet_id = sheet_row[0]

        # 2. Загружаем "сырые" данные
        cursor.execute(
            "SELECT cell_address, value FROM raw_data WHERE sheet_name = ? ORDER BY cell_address",
            (sheet_name,)
        )
        raw_data_rows = cursor.fetchall()

        # 3. Загружаем стили
        cursor.execute(
            "SELECT range_address, style_attributes FROM sheet_styles WHERE sheet_id = ?",
            (sheet_id,)
        )
        style_rows = cursor.fetchall()

        conn.close()

        # --- Обработка сырых данных ---
        max_row = -1
        max_col = -1
        data_map = {}

        for addr, val in raw_data_rows:
            parsed = _parse_cell_address(addr)
            if parsed is None:
                logger.warning(f"fetch_sheet_data: Не удалось распарсить адрес ячейки '{addr}' для листа '{sheet_name}'")
                continue
            row_idx, col_idx = parsed
            data_map[(row_idx, col_idx)] = val
            max_row = max(max_row, row_idx)
            max_col = max(max_col, col_idx)

        # Создание 2D списка
        if max_row >= 0 and max_col >= 0:
            rows_2d = [[""] * (max_col + 1) for _ in range(max_row + 1)]
            for (r, c), val in data_map.items():
                if 0 <= r <= max_row and 0 <= c <= max_col:
                    rows_2d[r][c] = val
        # --------------------------

        # --- Обработка стилей ---
        for range_addr, style_attrs_json in style_rows:
            parsed_range = _parse_range_address(range_addr)
            if parsed_range is None:
                logger.warning(f"fetch_sheet_data: Не удалось распарсить адрес диапазона '{range_addr}' для листа '{sheet_name}'")
                continue
            start_r, start_c, end_r, end_c = parsed_range

            try:
                style_attrs = json.loads(style_attrs_json) if isinstance(style_attrs_json, str) else style_attrs_json
                if not isinstance(style_attrs, dict):
                     logger.warning(f"fetch_sheet_data: 'style_attributes' для '{range_addr}' не является словарем или корректным JSON-объектом. Пропущено.")
                     continue
            except json.JSONDecodeError as e:
                logger.warning(f"fetch_sheet_data: Ошибка парсинга JSON стиля для '{range_addr}': {e}")
                continue # Пропускаем некорректный стиль

            # Применяем стиль ко всем ячейкам в диапазоне
            for r in range(start_r, end_r + 1):
                for c in range(start_c, end_c + 1):
                    # Логируем загруженные стили для отладки (опционально, можно убрать)
                    # logger.debug(f"fetch_sheet_data: Установлен стиль для [{r}, {c}]: {list(style_attrs.keys())}")
                    styles_map[(r, c)] = style_attrs

        logger.info(f"fetch_sheet_data: Загружено {len(data_map)} ячеек данных и {len(styles_map)} стилей для листа '{sheet_name}'.")
        # ----------------------

    except Exception as e:
        logger.error(f"fetch_sheet_data: Ошибка при загрузке данных для листа '{sheet_name}' из БД {db_path}: {e}", exc_info=True)
        # Возвращаем пустые структуры в случае ошибки
        return [], {}

    return rows_2d, styles_map