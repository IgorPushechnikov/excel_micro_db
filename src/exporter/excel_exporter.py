# src/exporter/excel_exporter.py
"""
Модуль для экспорта данных проекта Excel Micro DB в новый Excel-файл с использованием XlsxWriter.
Экспортирует данные, формулы, стили и объединенные ячейки.
"""

import xlsxwriter
import logging
import sqlite3
from typing import Dict, Any, List, Optional, Tuple
from pathlib import Path
import re
import sys

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

# --- Вспомогательные функции ---

def _parse_cell_address_to_indices(address: str) -> Tuple[int, int]:
    """
    Преобразует адрес ячейки Excel (например, "A1") в (строка (0-based), столбец (0-based)).
    """
    match = re.match(r"([A-Z]+)(\d+)", address)
    if not match:
        logger.warning(f"Невозможно распарсить адрес ячейки: {address}")
        return (-1, -1)
    column_letter, row_number = match.groups()
    row = int(row_number) - 1
    col = 0
    for char in column_letter:
        col = col * 26 + (ord(char) - ord('A') + 1)
    col = col - 1 # 0-based
    return (row, col)

def _parse_range_address_to_indices(range_address: str) -> Tuple[int, int, int, int]:
    """
    Преобразует адрес диапазона Excel (например, "A1:B2" или "C5")
    в (start_row (0-based), start_col (0-based), end_row (0-based), end_col (0-based)).
    """
    if ':' in range_address:
        start_addr, end_addr = range_address.split(':')
    else:
        start_addr = end_addr = range_address

    start_row, start_col = _parse_cell_address_to_indices(start_addr)
    end_row, end_col = _parse_cell_address_to_indices(end_addr)

    if start_row == -1 or end_row == -1:
        logger.error(f"Ошибка при парсинге диапазона: {range_address}")
        return (0, 0, -1, -1)

    return (start_row, start_col, end_row, end_col)

def _convert_style_attributes_to_xlsxwriter_format_dict(style_attributes: Dict[str, Any]) -> Dict[str, Any]:
    """
    Преобразует атрибуты стиля из формата БД в словарь атрибутов для workbook.add_format().
    """
    format_props = {}

    # --- Шрифт ---
    if 'font_name' in style_attributes and style_attributes['font_name'] is not None:
        format_props['font_name'] = style_attributes['font_name']
    if 'font_sz' in style_attributes and style_attributes['font_sz'] is not None:
        format_props['font_size'] = float(style_attributes['font_sz'])
    if 'font_b' in style_attributes:
        format_props['bold'] = bool(style_attributes['font_b'])
    if 'font_i' in style_attributes:
        format_props['italic'] = bool(style_attributes['font_i'])
    if 'font_u' in style_attributes and style_attributes['font_u'] is not None:
        format_props['underline'] = style_attributes['font_u']
    if 'font_strike' in style_attributes:
        format_props['font_strikeout'] = bool(style_attributes['font_strike'])
    if 'font_color_rgb' in style_attributes and style_attributes['font_color_rgb'] is not None:
        color_val = style_attributes['font_color_rgb']
        if not color_val.startswith('#'):
            if len(color_val) == 6 or len(color_val) == 8:
                 format_props['font_color'] = f"#{color_val[-6:]}"
            else:
                 logger.warning(f"Неожиданный формат цвета шрифта: {color_val}")
        else:
             format_props['font_color'] = color_val

    # --- Заливка ---
    pattern_type = style_attributes.get('fill_pattern_type')
    if pattern_type:
        format_props['pattern'] = 1 if pattern_type == 'solid' else 0
        if 'fill_fg_color_rgb' in style_attributes and style_attributes['fill_fg_color_rgb'] is not None:
            color_val = style_attributes['fill_fg_color_rgb']
            if not color_val.startswith('#'):
                if len(color_val) == 6 or len(color_val) == 8:
                     format_props['bg_color'] = f"#{color_val[-6:]}"
                else:
                     logger.warning(f"Неожиданный формат цвета заливки: {color_val}")
            else:
                 format_props['bg_color'] = color_val

    # --- Границы ---
    border_props = {}
    for side in ['left', 'right', 'top', 'bottom']:
        side_style_key = f'border_{side}_style'
        side_color_key = f'border_{side}_color_rgb'
        side_style = style_attributes.get(side_style_key)
        if side_style:
            border_props[side] = {'style': side_style}
            if side_color_key in style_attributes and style_attributes[side_color_key] is not None:
                color_val = style_attributes[side_color_key]
                if not color_val.startswith('#'):
                    if len(color_val) == 6 or len(color_val) == 8:
                         border_props[side]['color'] = f"#{color_val[-6:]}"
                    else:
                         logger.warning(f"Неожиданный формат цвета границы ({side}): {color_val}")
                else:
                     border_props[side]['color'] = color_val

    diag_up = style_attributes.get('border_diagonal_up')
    diag_down = style_attributes.get('border_diagonal_down')
    diag_type = 0
    if diag_up and diag_down:
        diag_type = 3
    elif diag_down:
        diag_type = 1
    elif diag_up:
        diag_type = 2
    if diag_type != 0:
        border_props['diag_type'] = diag_type
        diag_style = style_attributes.get('border_diagonal_style')
        if diag_style:
            border_props['diag_border'] = {'style': diag_style}
            if 'border_diagonal_color_rgb' in style_attributes and style_attributes['border_diagonal_color_rgb'] is not None:
                color_val = style_attributes['border_diagonal_color_rgb']
                if not color_val.startswith('#'):
                    if len(color_val) == 6 or len(color_val) == 8:
                         border_props['diag_border']['color'] = f"#{color_val[-6:]}"
                    else:
                         logger.warning(f"Неожиданный формат цвета диагональной границы: {color_val}")
                else:
                     border_props['diag_border']['color'] = color_val

    if border_props:
        format_props.update(border_props)

    # --- Выравнивание ---
    h_align = style_attributes.get('alignment_horizontal')
    if h_align:
        format_props['align'] = h_align
    v_align = style_attributes.get('alignment_vertical')
    if v_align:
        format_props['valign'] = v_align
    if 'alignment_wrap_text' in style_attributes:
        format_props['text_wrap'] = bool(style_attributes['alignment_wrap_text'])
    if 'alignment_shrink_to_fit' in style_attributes:
        format_props['shrink'] = bool(style_attributes['alignment_shrink_to_fit'])

    # --- Защита ---
    if 'protection_locked' in style_attributes:
        format_props['locked'] = bool(style_attributes['protection_locked'])
    if 'protection_hidden' in style_attributes:
        format_props['hidden'] = bool(style_attributes['protection_hidden'])

    # TODO: num_format - требуется маппинг ID -> строка формата
    # if 'num_fmt_id' in style_attributes and style_attributes['num_fmt_id'] is not None:
    #     format_props['num_format'] = ... 

    logger.debug(f"Преобразованы атрибуты стиля: {format_props}")
    return format_props

# --- Основные функции экспорта ---

def export_project_from_db(db_path: str, output_path: str) -> bool:
    """
    Экспортирует проект из SQLite БД в файл Excel (.xlsx) с использованием XlsxWriter.
    Загружает данные для каждого листа напрямую из storage.

    Args:
        db_path (str): Путь к файлу БД проекта (.sqlite).
        output_path (str): Путь к файлу Excel, который будет создан.

    Returns:
        bool: True, если экспорт прошёл успешно, иначе False.
    """
    logger.info("=== НАЧАЛО ЭКСПОРТА ПРОЕКТА ИЗ БД (XlsxWriter) ===")
    logger.info(f"Путь к БД проекта: {db_path}")
    logger.info(f"Путь к выходному файлу: {output_path}")

    db_path_obj = Path(db_path)
    output_path_obj = Path(output_path)

    if not db_path_obj.exists():
        logger.error(f"Файл БД проекта не найден: {db_path}")
        return False

    # Инициализация workbook вне блока try для корректной обработки в except
    workbook = None
    try:
        # Импорт внутри блока try, чтобы избежать проблем при импорте модуля
        from src.storage.base import ProjectDBStorage
        
        workbook = xlsxwriter.Workbook(str(output_path_obj))
        logger.info("Создана новая книга Excel (XlsxWriter).")

        # --- Получаем список листов напрямую из БД ---
        sheet_list = []
        db_conn = None
        try:
            db_conn = sqlite3.connect(str(db_path_obj))
            db_conn.row_factory = sqlite3.Row
            cursor = db_conn.cursor()
            cursor.execute("SELECT id, name FROM sheets ORDER BY sheet_index")
            sheet_rows = cursor.fetchall()
            sheet_list = [(row['id'], row['name']) for row in sheet_rows]
            logger.info(f"Найдено {len(sheet_list)} листов для экспорта.")
        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при получении списка листов: {e}")
            raise
        except Exception as e:
            logger.error(f"Неожиданная ошибка при получении списка листов: {e}")
            raise
        finally:
            if db_conn:
                db_conn.close()
        # ---------------------------------------------

        if not sheet_list:
             logger.warning("В проекте не найдено листов. Создается пустой файл.")
             workbook.add_worksheet("EmptySheet")
             workbook.close()
             logger.info(f"Пустой файл сохранен: {output_path}")
             return True

        # --- Экспорт каждого листа ---
        with ProjectDBStorage(str(db_path_obj)) as storage:
            logger.debug("Подключение к БД проекта установлено через ProjectDBStorage.")
            for sheet_id, sheet_name in sheet_list:
                logger.info(f"Экспорт листа: {sheet_name} (ID: {sheet_id})")
                worksheet = workbook.add_worksheet(sheet_name)
                # Передаем экземпляр storage и имя/ID листа
                _export_sheet_content_with_styles(workbook, worksheet, storage, sheet_id, sheet_name)
                # Объединенные ячейки также экспортируем отдельно, если нужно
                # (или интегрировать в _export_sheet_content_with_styles)
                # Для простоты оставим отдельно, как в оригинале каркаса
                # _export_sheet_merged_cells(worksheet, storage, sheet_id, sheet_name) 
                # TODO: Реализовать _export_sheet_merged_cells, если они не экспортируются в _export_sheet_content_with_styles

        workbook.close()
        logger.info(f"Файл Excel успешно сохранен: {output_path}")
        logger.info("=== ЭКСПОРТ ПРОЕКТА ИЗ БД (XlsxWriter) ЗАВЕРШЕН ===")
        return True

    except Exception as e:
        logger.error(f"Ошибка при экспорте проекта в файл '{output_path}': {e}", exc_info=True)
        # Пытаемся закрыть workbook, если он был создан, чтобы избежать повреждения файла
        try:
            if workbook is not None:
                workbook.close()
        except Exception as close_error:
             logger.error(f"Ошибка при закрытии книги Excel: {close_error}")
        return False

def _export_sheet_content_with_styles(workbook, worksheet, storage, sheet_id: int, sheet_name: str) -> None:
    """
    Экспортирует данные и формулы листа, применяя стили одновременно.
    Загружает данные напрямую через экземпляр storage.

    Args:
        workbook: Экземпляр xlsxwriter.Workbook.
        worksheet: Экземпляр xlsxwriter.Workbook.add_worksheet.
        storage: Экземпляр ProjectDBStorage для загрузки данных.
        sheet_id (int): ID листа в БД.
        sheet_name (str): Имя листа.
    """
    try:
        logger.debug(f"Начало экспорта содержимого листа '{sheet_name}' с применением стилей.")

        # --- 1. Загрузка данных ---
        logger.debug(f"Загрузка редактируемых данных для листа '{sheet_name}'...")
        editable_data_result = storage.load_sheet_editable_data(sheet_name)
        
        # === ДОБАВЛЕНО: Расширенное логирование результата загрузки данных ===
        logger.debug(f"Результат load_sheet_editable_data для '{sheet_name}': {type(editable_data_result)}")
        if isinstance(editable_data_result, dict):
            logger.debug(f"  - Ключи в результате: {list(editable_data_result.keys())}")
            logger.debug(f"  - column_names: {editable_data_result.get('column_names', 'N/A')}")
            rows_data_log = editable_data_result.get('rows', [])
            logger.debug(f"  - Количество строк данных: {len(rows_data_log)}")
            if rows_data_log:
                logger.debug(f"  - Пример первой строки: {rows_data_log[0] if len(rows_data_log) > 0 else 'N/A'}")
                logger.debug(f"  - Тип первой строки: {type(rows_data_log[0]) if len(rows_data_log) > 0 else 'N/A'}")
        # ===================================================================
        
        # === ИСПРАВЛЕНО: Проверка типа результата ===
        if not isinstance(editable_data_result, dict):
            logger.error(f"load_sheet_editable_data для листа '{sheet_name}' вернула {type(editable_data_result)}, ожидался dict.")
            editable_data_result = {"column_names": [], "rows": []}
        
        column_names = editable_data_result.get("column_names", [])
        # === ИСПРАВЛЕНО: Обработка rows как списка кортежей ===
        rows_as_tuples = editable_data_result.get("rows", []) 
        # ================================

        if not column_names:
            logger.warning(f"Нет данных (column_names пуст) для экспорта на листе '{sheet_name}'.")
            return # Возвращаемся, если данных нет

        # Промежуточная структура: {(row, col): (value, format_dict)}
        sheet_content: Dict[Tuple[int, int], Tuple[Any, Optional[Dict[str, Any]]]] = {}

        # --- 2. Запись заголовков (строка 0) ---
        for col_idx, col_name in enumerate(column_names):
            sheet_content[(0, col_idx)] = (col_name, None)

        # --- 3. Запись данных (начиная со строки 1) ---
        # === ИСПРАВЛЕНО: Обработка rows как списка кортежей ===
        for row_idx, row_tuple in enumerate(rows_as_tuples, start=1):
            for col_idx, value in enumerate(row_tuple):
                 if col_idx < len(column_names):
                     sheet_content[(row_idx, col_idx)] = (value, None)
                 else:
                     logger.warning(f"Строка {row_idx} содержит больше значений, чем ожидаемых столбцов. Лишние значения проигнорированы.")
        # ================================
        logger.debug(f"Размер sheet_content после загрузки данных: {len(sheet_content)}")

        # --- 4. Загрузка и применение формул ---
        logger.debug(f"Загрузка формул для листа '{sheet_name}' (ID: {sheet_id})...")
        formulas_data = storage.load_sheet_formulas(sheet_id) # <-- ИСПРАВЛЕНО: правильное имя переменной
        logger.debug(f"Найдено {len(formulas_data)} формул для экспорта на листе '{sheet_name}'.")
        # === ДОБАВЛЕНО: Расширенное логирование формул ===
        if formulas_data:
            logger.debug(f"  - Пример первой формулы: {formulas_data[0]}")
        # =============================================
        for formula_info in formulas_data: # <-- ИСПРАВЛЕНО: правильное имя переменной
            cell_address = formula_info.get("cell", "")
            formula = formula_info.get("formula", "")
            if cell_address and formula:
                row_idx, col_idx = _parse_cell_address_to_indices(cell_address)
                if row_idx != -1 and col_idx != -1:
                    formula_to_write = formula if formula.startswith('=') else f"={formula}"
                    sheet_content[(row_idx, col_idx)] = (formula_to_write, None)
                    logger.debug(f"Формула добавлена для ячейки ({row_idx}, {col_idx}): {formula_to_write}")
                else:
                     logger.warning(f"Не удалось распарсить адрес формулы: {cell_address}")
        logger.debug(f"Размер sheet_content после загрузки формул: {len(sheet_content)}")

        # --- 5. Загрузка и применение стилей ---
        logger.debug(f"Загрузка стилей для листа '{sheet_name}' (ID: {sheet_id})...")
        styled_ranges_data = storage.load_sheet_styles(sheet_id) # <-- ИСПРАВЛЕНО: правильное имя переменной
        logger.debug(f"Применение {len(styled_ranges_data)} стилевых диапазонов на листе '{sheet_name}'.")
        # === ДОБАВЛЕНО: Расширенное логирование стилей ===
        if styled_ranges_data:
            logger.debug(f"  - Пример первого стиля: {styled_ranges_data[0]}")
        # =============================================
        
        format_cache: Dict[str, Any] = {}

        for style_range_info in styled_ranges_data: # <-- ИСПРАВЛЕНО: правильное имя переменной
            range_address = style_range_info.get("range_address")
            style_attributes = style_range_info.get("style_attributes", {})

            if not range_address or not style_attributes:
                continue

            format_dict = _convert_style_attributes_to_xlsxwriter_format_dict(style_attributes)
            
            if not format_dict:
                logger.debug(f"Преобразование стиля для диапазона {range_address} не дало параметров.")
                continue

            cache_key = str(sorted(format_dict.items()))
            if cache_key in format_cache:
                cell_format = format_cache[cache_key]
                logger.debug(f"Формат для стиля из кэша: {cache_key}")
            else:
                try:
                    cell_format = workbook.add_format(format_dict)
                    format_cache[cache_key] = cell_format
                    logger.debug(f"Создан новый формат для стиля: {cache_key}")
                except Exception as e:
                    logger.error(f"Ошибка создания формата XlsxWriter: {e}")
                    continue

            start_row, start_col, end_row, end_col = _parse_range_address_to_indices(range_address)
            if end_row < start_row or end_col < start_col:
                logger.warning(f"Некорректный диапазон стиля: {range_address}")
                continue

            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    current_value, _ = sheet_content.get((r, c), ("", None))
                    sheet_content[(r, c)] = (current_value, format_dict)
                    logger.debug(f"Применен стиль к ячейке ({r}, {c})")

        logger.debug(f"Размер sheet_content после загрузки стилей: {len(sheet_content)}")
        # --- 6. Запись в worksheet из промежуточной структуры ---
        logger.debug(f"Запись содержимого листа '{sheet_name}' в файл Excel. Всего ячеек: {len(sheet_content)}")
        for (row, col), (value, format_dict) in sheet_content.items():
            format_to_use = None
            if format_dict:
                 cache_key = str(sorted(format_dict.items()))
                 format_to_use = format_cache.get(cache_key)
            
            if isinstance(value, str) and value.startswith('='):
                 worksheet.write_formula(row, col, value, format_to_use)
                 logger.debug(f"Записана формула в ({row}, {col}): {value}")
            else:
                 worksheet.write(row, col, value, format_to_use)
                 logger.debug(f"Записано значение в ({row}, {col}): {value}")

        logger.debug(f"Экспорт содержимого листа '{sheet_name}' с применением стилей завершен.")

    except Exception as e:
        logger.error(f"Ошибка при экспорте содержимого/стилей листа '{sheet_name}': {e}", exc_info=True)
        # Можно выбрасывать исключение выше или продолжать с другими листами
        # raise # Выбрасываем, чтобы ошибка поднялась в export_project_from_db

# Точка входа для тестирования напрямую
if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description="Экспорт проекта Excel Micro DB напрямую из БД с использованием XlsxWriter.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("db_path", help="Путь к файлу project_data.db")
    parser.add_argument("output_path", help="Путь для сохранения выходного .xlsx файла")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
                        help="Уровень логгирования (по умолчанию INFO)")

    args = parser.parse_args()

    # Настройка логирования для прямого запуска
    log_level = getattr(logging, args.log_level.upper())
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_formatter = logging.Formatter(log_format)
    logger.handlers.clear()
    logger.addHandler(console_handler)
    logger.setLevel(log_level)

    logger.info("=== ЗАПУСК СКРИПТА ЭКСПОРТА (XlsxWriter) ===")

    success = export_project_from_db(args.db_path, args.output_path)

    if success:
        logger.info(f"Экспорт успешно завершен. Файл сохранен в: {args.output_path}")
        sys.exit(0)
    else:
        logger.error(f"Экспорт завершился с ошибкой.")
        sys.exit(1)
