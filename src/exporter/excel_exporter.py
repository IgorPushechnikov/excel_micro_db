# src/exporter/excel_exporter.py
"""
Модуль для экспорта данных проекта Excel Micro DB в новый Excel-файл с использованием XlsxWriter.
Экспортирует данные, формулы, стили и объединенные ячейки.
"""

import xlsxwriter
import logging
from typing import Dict, Any, List, Optional, Tuple # Импортируем Any
from pathlib import Path
import re # Для парсинга адресов диапазонов
import sys

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

# Импорт storage должен быть внутри функции или в блоке main,
# чтобы избежать циклических импортов, если storage когда-нибудь импортирует этот модуль.
# from src.storage.base import ProjectDBStorage

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
        # 'solid' в openpyxl -> 1 в XlsxWriter, другие -> 0 или нужно сопоставлять
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
            # В XlsxWriter граница определяется словарем {'style': ..., 'color': ...}
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

    # Диагональные границы
    diag_up = style_attributes.get('border_diagonal_up')
    diag_down = style_attributes.get('border_diagonal_down')
    diag_type = 0 # 0=none, 1=down, 2=up, 3=both
    if diag_up and diag_down:
        diag_type = 3 # both
    elif diag_down:
        diag_type = 1 # down
    elif diag_up:
        diag_type = 2 # up
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
        # Некоторые значения могут отличаться, например 'centerContinuous' -> 'center_across'
        # Пока используем напрямую, можно добавить сопоставление при необходимости.
        format_props['align'] = h_align # 'left', 'center', 'right' и т.д. должны совпадать
    v_align = style_attributes.get('alignment_vertical')
    if v_align:
        # 'top', 'center', 'bottom' должны совпадать
        format_props['valign'] = v_align
    if 'alignment_wrap_text' in style_attributes:
        format_props['text_wrap'] = bool(style_attributes['alignment_wrap_text'])
    if 'alignment_shrink_to_fit' in style_attributes:
        # XlsxWriter использует 'shrink'
        format_props['shrink'] = bool(style_attributes['alignment_shrink_to_fit'])
    # TODO: Обработка других атрибутов выравнивания (indent, rotation, reading_order и т.д.)
    # если они будут использоваться.

    # --- Защита ---
    # XlsxWriter использует 'locked' и 'hidden' напрямую в формате
    if 'protection_locked' in style_attributes:
        format_props['locked'] = bool(style_attributes['protection_locked'])
    if 'protection_hidden' in style_attributes:
        format_props['hidden'] = bool(style_attributes['protection_hidden'])

    # --- Другие атрибуты ---
    # num_format - формат чисел
    # XlsxWriter использует 'num_format'
    # if 'num_fmt_id' in style_attributes and style_attributes['num_fmt_id'] is not None:
    #     # TODO: Нужно сопоставить num_fmt_id с реальным строковым форматом.
    #     # Пока оставим заглушку или пропустим, если это не критично на данном этапе.
    #     # format_props['num_format'] = ...
    #     pass # Пропускаем, так как у нас нет маппинга ID -> строка формата

    logger.debug(f"Преобразованы атрибуты стиля: {format_props}")
    return format_props

# --- Основные функции экспорта ---

def export_project_from_db(db_path: str, output_path: str) -> bool:
    """
    Экспортирует проект из SQLite БД в файл Excel (.xlsx) с использованием XlsxWriter.

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

    project_data = {}
    # === ИСПРАВЛЕНО: Инициализация workbook для устранения ошибки Pylance ===
    workbook = None
    # ========================================================================
    try:
        # Импорт внутри блока try, чтобы избежать проблем при импорте модуля
        from src.storage.base import ProjectDBStorage
        with ProjectDBStorage(str(db_path_obj)) as storage:
            logger.debug("Подключение к БД проекта установлено.")
            project_data = storage.get_all_data()
            if not project_data:
                 logger.error("Не удалось загрузить данные проекта из БД.")
                 return False
            logger.debug("Данные проекта успешно загружены из БД.")
    except Exception as e:
        logger.error(f"Ошибка при подключении к БД или загрузке данных: {e}", exc_info=True)
        return False

    try:
        workbook = xlsxwriter.Workbook(str(output_path_obj))
        logger.info("Создана новая книга Excel (XlsxWriter).")

        project_info = project_data.get("project_info", {})
        project_name = project_info.get("name", "Unknown Project")
        sheets_data = project_data.get("sheets", {})

        logger.info(f"Найдено {len(sheets_data)} листов для экспорта.")

        if not sheets_data:
             logger.warning("В проекте не найдено листов. Создается пустой файл.")
             workbook.add_worksheet("EmptySheet")
             workbook.close()
             logger.info(f"Пустой файл сохранен: {output_path}")
             return True

        for sheet_name, sheet_info in sheets_data.items():
            logger.info(f"Экспорт листа: {sheet_name}")
            worksheet = workbook.add_worksheet(sheet_name)
            _export_sheet_content_with_styles(workbook, worksheet, sheet_info)
            _export_sheet_merged_cells(worksheet, sheet_info)

        workbook.close()
        logger.info(f"Файл Excel успешно сохранен: {output_path}")
        logger.info("=== ЭКСПОРТ ПРОЕКТА ИЗ БД (XlsxWriter) ЗАВЕРШЕН ===")
        return True

    except Exception as e:
        logger.error(f"Ошибка при экспорте проекта в файл '{output_path}': {e}", exc_info=True)
        # === ИСПРАВЛЕНО: Проверка перед закрытием workbook ===
        # Пытаемся закрыть workbook, если он был создан, чтобы избежать повреждения файла
        try:
            if workbook is not None: # Проверяем, был ли workbook инициализирован
                workbook.close()
        except Exception as close_error:
             logger.error(f"Ошибка при закрытии книги Excel: {close_error}")
        # =====================================================
        return False

def _export_sheet_content_with_styles(workbook, worksheet, sheet_info: Dict[str, Any]) -> None:
    """
    Экспортирует данные и формулы листа, применяя стили одновременно.
    Собирает данные и стили в промежуточную структуру, затем записывает.
    """
    try:
        logger.debug("Начало экспорта содержимого листа с применением стилей.")

        # --- 1. Сбор данных ---
        # ВАЖНО: Убедиться, что данные берутся из правильного источника.
        # Согласно документации и анализу, для экспорта результатов нужно использовать editable_data.
        editable_data = sheet_info.get("editable_data", {}) # Предполагается, что это загружено storage.load_sheet_editable_data
        column_names = editable_data.get("column_names", [])
        rows = editable_data.get("rows", [])

        if not column_names:
            logger.warning("Нет данных для экспорта на листе.")
            return

        # Промежуточная структура: {(row, col): (value, format_dict)}
        # row, col - 0-based индексы
        # format_dict хранится для возможности кэширования объекта формата позже
        sheet_content: Dict[Tuple[int, int], Tuple[Any, Optional[Dict[str, Any]]]] = {}

        # Записываем заголовки (строка 0)
        for col_idx, col_name in enumerate(column_names):
            sheet_content[(0, col_idx)] = (col_name, None) # Заголовки без стиля из БД

        # Записываем данные (начиная со строки 1)
        for row_idx, row_data in enumerate(rows, start=1):
            for col_idx, col_name in enumerate(column_names):
                value = row_data.get(col_name, "")
                sheet_content[(row_idx, col_idx)] = (value, None) # Пока без стиля

        # --- 2. Сбор формул ---
        # Формулы могут быть в любой строке/столбце, включая заголовки
        formulas_data = sheet_info.get("formulas", [])
        logger.debug(f"Найдено {len(formulas_data)} формул для экспорта.")
        for formula_info in formulas_data:
            cell_address = formula_info.get("cell", "")  # Например, "F2"
            formula = formula_info.get("formula", "")    # Например, "=SUM(B2:E2)"

            if cell_address and formula:
                # Преобразуем адрес ячейки в индексы
                row_idx, col_idx = _parse_cell_address_to_indices(cell_address)
                if row_idx != -1 and col_idx != -1:
                    # XlsxWriter позволяет записывать формулы напрямую
                    # Убедиться, что формула начинается с '='
                    # Формулы в БД могут храниться как с '=', так и без
                    formula_to_write = formula if formula.startswith('=') else f"={formula}"
                    # Формула заменяет значение. Стиль будет применен позже.
                    sheet_content[(row_idx, col_idx)] = (formula_to_write, None) # Пока без стиля
                    logger.debug(f"Формула добавлена для ячейки ({row_idx}, {col_idx}): {formula_to_write}")
                else:
                     logger.warning(f"Не удалось распарсить адрес формулы: {cell_address}")

        # --- 3. Сбор стилей и применение к содержимому ---
        # Согласно анализу storage/styles.py, load_sheet_styles возвращает
        # List[Dict[str, Any]] с ключами 'style_attributes' и 'range_address'
        styled_ranges_data = sheet_info.get("styled_ranges", [])
        logger.debug(f"Применение {len(styled_ranges_data)} стилевых диапазонов.")

        # Кэш для преобразованных форматов, чтобы не создавать одинаковые объекты
        # Используем Dict[str, Any] для хранения объектов формата XlsxWriter
        # === ИСПРАВЛЕНО: Тип кэша ===
        format_cache: Dict[str, Any] = {} # workbook.add_format(...) возвращает объект формата
        # =========================

        for style_range_info in styled_ranges_data:
            range_address = style_range_info.get("range_address")
            # Стиль содержится в ключе 'style_attributes'
            style_attributes = style_range_info.get("style_attributes", {})

            if not range_address:
                logger.warning("Пропущен стиль из-за отсутствия range_address.")
                continue
            if not style_attributes:
                logger.debug(f"Пропущен стиль для диапазона {range_address} из-за отсутствия атрибутов.")
                continue # Нечего применять

            # Преобразуем атрибуты стиля из формата БД в словарь для add_format
            # Этот словарь будет использоваться как ключ для кэша
            format_dict = _convert_style_attributes_to_xlsxwriter_format_dict(style_attributes)
            
            if not format_dict:
                logger.debug(f"Преобразование стиля для диапазона {range_address} не дало параметров формата. Пропущено.")
                continue # Нечего применять

            # Создаем уникальный ключ для кэша на основе словаря атрибутов
            # sorted + str для детерминированного ключа
            # === ИСПРАВЛЕНО: Создание ключа кэша ===
            cache_key = str(sorted(format_dict.items()))
            # =====================================

            # Проверяем кэш
            if cache_key in format_cache:
                cell_format = format_cache[cache_key]
                logger.debug(f"Формат для стиля из кэша: {cache_key}")
            else:
                # Создаем новый объект формата XlsxWriter
                try:
                    cell_format = workbook.add_format(format_dict)
                    # Сохраняем в кэш
                    format_cache[cache_key] = cell_format
                    logger.debug(f"Создан новый формат для стиля: {cache_key}")
                except Exception as format_error:
                    logger.error(f"Ошибка при создании формата XlsxWriter для диапазона {range_address} с атрибутами {format_dict}: {format_error}")
                    continue # Пропускаем этот стиль

            # Парсим адрес диапазона
            start_row, start_col, end_row, end_col = _parse_range_address_to_indices(range_address)
            
            # Проверка на корректность диапазона
            if end_row < start_row or end_col < start_col:
                logger.warning(f"Некорректный диапазон для стиля: {range_address}. Пропущено.")
                continue
                
            logger.debug(f"Применение стиля к диапазону {range_address} (строки {start_row}-{end_row}, столбцы {start_col}-{end_col}).")

            # Итерируемся по строкам и столбцам диапазона
            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    # Получаем текущее значение ячейки из промежуточной структуры
                    current_value, _ = sheet_content.get((r, c), ("", None))
                    # Обновляем запись в промежуточной структуре с форматом
                    # Сохраняем dict для лога и кэширования, сам объект формата в кэше
                    sheet_content[(r, c)] = (current_value, format_dict) 
                    logger.debug(f"Применен стиль к ячейке ({r}, {c})")

        # --- 4. Запись в worksheet из промежуточной структуры ---
        logger.debug("Запись содержимого листа в файл Excel.")
        for (row, col), (value, format_dict) in sheet_content.items():
            # Получаем объект формата из кэша, если он был применен
            format_to_use = None
            if format_dict:
                 # Воссоздаем ключ кэша
                 cache_key = str(sorted(format_dict.items()))
                 format_to_use = format_cache.get(cache_key) # Может вернуть None, если ключ не найден (хотя не должно)
            
            # XlsxWriter различает запись данных и формул
            if isinstance(value, str) and value.startswith('='):
                 # Это формула
                 worksheet.write_formula(row, col, value, format_to_use)
                 logger.debug(f"Записана формула в ({row}, {col}): {value}")
            else:
                 # Это значение (данные или заголовок)
                 worksheet.write(row, col, value, format_to_use)
                 logger.debug(f"Записано значение в ({row}, {col}): {value}")

        logger.debug("Экспорт содержимого листа с применением стилей завершен.")

    except Exception as e:
        logger.error(f"Ошибка при экспорте содержимого/стилей листа: {e}", exc_info=True)

def _export_sheet_merged_cells(worksheet, sheet_info: Dict[str, Any]) -> None:
    """Экспортирует объединенные ячейки листа."""
    try:
        # Объединенные ячейки хранятся в ключе 'merged_cells'
        merged_cells_data = sheet_info.get("merged_cells", [])
        logger.debug(f"Экспорт {len(merged_cells_data)} объединенных диапазонов.")

        for range_address in merged_cells_data:
            if range_address:
                # XlsxWriter использует метод merge_range для объединения ячеек
                # Нужно определить значение для объединенного диапазона
                # Пока просто объединяем без значения. В будущем можно определить значение из данных.
                # merge_range(first_row, first_col, last_row, last_col, data, cell_format)
                # Нужно распарсить range_address
                start_row, start_col, end_row, end_col = _parse_range_address_to_indices(range_address)
                if end_row >= start_row and end_col >= start_col:
                    # Пишем пустую строку в объединенный диапазон
                    # TODO: Определить значение для merge_range (например, из левой верхней ячейки)
                    worksheet.merge_range(start_row, start_col, end_row, end_col, "", None)
                    logger.debug(f"Объединен диапазон: {range_address}")
                else:
                    logger.warning(f"Некорректный адрес объединенного диапазона: {range_address}")
    except Exception as e:
        logger.error(f"Ошибка при экспорте объединенных ячеек листа: {e}")

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
