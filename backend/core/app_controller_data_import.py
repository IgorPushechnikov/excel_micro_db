# backend/core/app_controller_data_import.py
"""
Модуль, содержащий функции для импорта различных типов данных из Excel-файла
в БД проекта через AppController.

Функции предназначены для вызова из AppController для реализации
импорта "по типам" (данные, стили, диаграммы, формулы) и "по режимам"
(всё, выборочно, частями).
"""

import logging
import os
from typing import Dict, Any, List, Optional, Union
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd
import numpy as np
# Исправлено: Импорт get_column_letter напрямую
from openpyxl.utils import get_column_letter

# Импортируем logger из utils
from backend.utils.logger import get_logger

# Импортируем функции из analyzer
# Исправлено: Импорт из правильного модуля и с правильными именами
from backend.analyzer.logic_documentation import _serialize_style, _serialize_chart

logger = get_logger(__name__)


# --- Вспомогательные функции ---

def _get_sheet_id_by_name(app_controller, sheet_name: str) -> Optional[int]:
    """
    Вспомогательная функция для получения sheet_id по имени листа.
    """
    if not app_controller.storage or not app_controller.storage.connection:
        return None

    try:
        cursor = app_controller.storage.connection.cursor()
        # Предполагаем project_id = 1
        cursor.execute("SELECT sheet_id FROM sheets WHERE name = ? AND project_id = 1", (sheet_name,))
        result = cursor.fetchone()
        return result[0] if result else None
    except Exception as e:
        logger.error(f"Ошибка при получении sheet_id для листа '{sheet_name}': {e}")
        return None


# --- Функции для импорта "всё" по типам ---

def import_raw_data_from_excel(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только "сырые" данные (значения ячеек) из Excel-файла в БД проекта.

    Args:
        app_controller: Экземпляр AppController.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not app_controller.storage:
        logger.error("Проект не загружен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта 'сырых' данных из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import = options.get('sheets', []) if options else []
        if not sheets_to_import:
            sheets_to_import = workbook.sheetnames

        for sheet_name in sheets_to_import:
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт 'сырых' данных с листа: {sheet_name}")
            sheet: Worksheet = workbook[sheet_name]

            raw_data_list = []
            for row in sheet.iter_rows(values_only=False):
                for cell in row:
                    if cell.value is not None or cell.data_type == 'f':
                        data_item = {
                            "cell_address": cell.coordinate,
                            "value": cell.value,
                        }
                        raw_data_list.append(data_item)

            if not app_controller.storage.save_sheet_raw_data(sheet_name, raw_data_list):
                logger.error(f"Не удалось сохранить 'сырые данные' для листа '{sheet_name}'.")
            else:
                logger.info(f"'Сырые данные' для листа '{sheet_name}' успешно импортированы.")

        logger.info(f"Импорт 'сырых' данных из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте 'сырых' данных из файла '{file_path}': {e}", exc_info=True)
        return False


def import_styles_from_excel(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только стили из Excel-файла в БД проекта.

    Args:
        app_controller: Экземпляр AppController.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not app_controller.storage:
        logger.error("Проект не загружен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта стилей из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import = options.get('sheets', []) if options else []
        if not sheets_to_import:
            sheets_to_import = workbook.sheetnames

        import json

        for sheet_name in sheets_to_import:
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт стилей с листа: {sheet_name}")
            sheet: Worksheet = workbook[sheet_name]

            style_ranges_map: Dict[str, List[str]] = {}

            for row in sheet.iter_rows(values_only=False):
                for cell in row:
                    style_dict = _serialize_style(cell)
                    if style_dict:
                        style_json = json.dumps(style_dict, sort_keys=True)
                        coord = cell.coordinate

                        if style_json in style_ranges_map:
                            style_ranges_map[style_json].append(coord)
                        else:
                            style_ranges_map[style_json] = [coord]

            styles_to_save = []
            for style_json, cell_addresses in style_ranges_map.items():
                for address in cell_addresses:
                    styles_to_save.append({
                        "range_address": address,
                        "style_attributes": style_json
                    })

            project_name = app_controller.current_project.get('name', 'Unknown') if app_controller.current_project else 'Unknown'
            sheet_id = _get_sheet_id_by_name(app_controller, sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось получить/создать ID для листа '{sheet_name}'. Пропущен.")
                continue

            if not app_controller.storage.save_sheet_styles(sheet_id, styles_to_save):
                logger.error(f"Не удалось сохранить стили для листа '{sheet_name}' (ID: {sheet_id}).")
            else:
                logger.info(f"Стили для листа '{sheet_name}' успешно импортированы.")

        logger.info(f"Импорт стилей из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте стилей из файла '{file_path}': {e}", exc_info=True)
        return False


def import_charts_from_excel(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только диаграммы из Excel-файла в БД проекта.

    Args:
        app_controller: Экземпляр AppController.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not app_controller.storage:
        logger.error("Проект не загружен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта диаграмм из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import = options.get('sheets', []) if options else []
        if not sheets_to_import:
            sheets_to_import = workbook.sheetnames

        for sheet_name in sheets_to_import:
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт диаграмм с листа: {sheet_name}")
            sheet: Worksheet = workbook[sheet_name]

            charts_list = []
            try:
                # Исправлено: Добавлен # type: ignore[attr-defined] для подавления ошибки Pylance
                charts_sheet = sheet._charts # type: ignore[attr-defined]
                for chart_obj in charts_sheet:
                    chart_data = _serialize_chart(chart_obj)
                    if chart_data:
                        charts_list.append({
                            "chart_data": chart_data
                        })
            except AttributeError as ae:
                logger.warning(f"Не удалось получить доступ к диаграммам листа '{sheet_name}' через _charts: {ae}")
            except Exception as e:
                logger.error(f"Ошибка при извлечении диаграмм с листа '{sheet_name}': {e}", exc_info=True)

            sheet_id = _get_sheet_id_by_name(app_controller, sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось получить/создать ID для листа '{sheet_name}'. Пропущен.")
                continue

            if not app_controller.storage.save_sheet_charts(sheet_id, charts_list):
                logger.error(f"Не удалось сохранить диаграммы для листа '{sheet_name}' (ID: {sheet_id}).")
            else:
                logger.info(f"Диаграммы для листа '{sheet_name}' успешно импортированы.")

        logger.info(f"Импорт диаграмм из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте диаграмм из файла '{file_path}': {e}", exc_info=True)
        return False


def import_formulas_from_excel(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только формулы из Excel-файла в БД проекта.

    Args:
        app_controller: Экземпляр AppController.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not app_controller.storage:
        logger.error("Проект не загружен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта формул из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import = options.get('sheets', []) if options else []
        if not sheets_to_import:
            sheets_to_import = workbook.sheetnames

        for sheet_name in sheets_to_import:
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт формул с листа: {sheet_name}")
            sheet: Worksheet = workbook[sheet_name]

            formulas_list = []
            for row in sheet.iter_rows(values_only=False):
                for cell in row:
                    # Исправлено: Проверяем, является ли значение строкой и начинается ли с '='
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                         formulas_list.append({
                             "cell_address": cell.coordinate,
                             "formula": cell.value # Сохраняем формулу как есть, включая '='
                         })

            sheet_id = _get_sheet_id_by_name(app_controller, sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось получить/создать ID для листа '{sheet_name}'. Пропущен.")
                continue

            if not app_controller.storage.save_sheet_formulas(sheet_id, formulas_list):
                logger.error(f"Не удалось сохранить формулы для листа '{sheet_name}' (ID: {sheet_id}).")
            else:
                logger.info(f"Формулы для листа '{sheet_name}' успешно импортированы.")

        logger.info(f"Импорт формул из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте формул из файла '{file_path}': {e}", exc_info=True)
        return False


# --- Функции для импорта "выборочно" по типам ---

# Заглушка для выборочного импорта. Реализация будет аналогична полному импорту,
# но с фильтрацией по листам/диапазонам, переданным в options.
# Например, options = {'sheets': ['Sheet1'], 'start_row': 1, 'end_row': 100}

def import_raw_data_from_excel_selective(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует "сырые" данные выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт сырых данных выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    # Реализация будет аналогична import_raw_data_from_excel, но с фильтрацией
    # на основе options (sheets, start_row, end_row, start_col, end_col)
    # Пока возвращаем True как успешное выполнение заглушки
    return True

def import_styles_from_excel_selective(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует стили выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт стилей выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    return True

def import_charts_from_excel_selective(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует диаграммы выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт диаграмм выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    return True

def import_formulas_from_excel_selective(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует формулы выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт формул выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    return True


# --- Функции для импорта "частями" по типам ---

# Заглушка для импорта частями. Реализация будет разбивать большой файл
# на части и вызывать соответствующую функцию импорта для каждой части.
# Например, импортировать по 1000 строк за раз.

def import_raw_data_from_excel_in_chunks(app_controller, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует "сырые" данные частями из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт сырых данных частями из '{file_path}' с опциями {chunk_options}. (Заглушка)")
    # Реализация будет разбивать файл на части и вызывать
    # import_raw_data_from_excel_selective для каждой части
    return True

def import_styles_from_excel_in_chunks(app_controller, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует стили частями из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт стилей частями из '{file_path}' с опциями {chunk_options}. (Заглушка)")
    return True

def import_charts_from_excel_in_chunks(app_controller, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует диаграммы частями из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт диаграмм частями из '{file_path}' с опциями {chunk_options}. (Заглушка)")
    return True

def import_formulas_from_excel_in_chunks(app_controller, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует формулы частями из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт формул частями из '{file_path}' с опциями {chunk_options}. (Заглушка)")
    return True


# --- Функция для БЫСТРОГО импорта "сырых" данных с помощью pandas ---

def import_raw_data_fast_with_pandas(app_controller, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Быстро импортирует "сырые" данные (значения ячеек) из Excel-файла с помощью pandas.

    Args:
        app_controller: Экземпляр AppController.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.
            {
                'sheets': List[str], # Список имен листов для импорта. Если пуст, все.
                'start_row': int,      # Начальная строка (0-based для pandas).
                'end_row': int,        # Конечная строка (0-based, не включительно для pandas).
                'start_col': int,      # Начальный столбец (0-based для pandas).
                'end_col': int,        # Конечный столбец (0-based, не включительно для pandas).
                'header': int or None,  # Номер строки для заголовков (0-based). None если нет заголовков.
                'dtype': dict,         # Типы данных для столбцов (например, {'A': str, 'B': float})
                'engine': str,         # Движок pandas ('xlrd', 'openpyxl', 'odf', 'pyxlsb'). По умолчанию 'openpyxl'.
            }

    Returns:
        bool: True, если импорт успешен, иначе False.
    """
    if not app_controller.storage:
        logger.error("Проект не загружен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало БЫСТРОГО импорта 'сырых' данных из Excel-файла: {file_path} с опциями {options}")

    try:
        # Установим значения по умолчанию для опций
        if options is None:
            options = {}
        sheets_to_import = options.get('sheets', None) # None означает все листы
        start_row_pandas = options.get('start_row', None) # 0-based для pandas
        end_row_pandas = options.get('end_row', None) # 0-based, не включительно
        start_col_pandas = options.get('start_col', None) # 0-based для pandas
        end_col_pandas = options.get('end_col', None) # 0-based, не включительно
        header_pandas = options.get('header', 0) # По умолчанию первая строка - заголовок
        dtype_dict = options.get('dtype', None)
        engine_pandas = options.get('engine', 'openpyxl') # По умолчанию openpyxl

        # Используем pd.ExcelFile для возможности чтения нескольких листов
        with pd.ExcelFile(file_path, engine=engine_pandas) as xls_file:
            # Определяем список листов для импорта
            sheets_to_process = sheets_to_import if sheets_to_import else xls_file.sheet_names

            for sheet_name in sheets_to_process:
                if sheet_name not in xls_file.sheet_names:
                    logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                    continue

                logger.info(f"Быстрый импорт данных с листа: {sheet_name}")

                # --- Чтение данных с помощью pandas ---
                df = pd.read_excel(
                    xls_file,
                    sheet_name=sheet_name,
                    header=header_pandas,
                    index_col=None, # Не использовать столбец как индекс
                    usecols=None if (start_col_pandas is None and end_col_pandas is None) else lambda x: (x >= (start_col_pandas or 0)) & (x < (end_col_pandas or float('inf'))),
                    skiprows=start_row_pandas,
                    nrows=None if end_row_pandas is None else (end_row_pandas - (start_row_pandas or 0)),
                    dtype=dtype_dict,
                    engine=None # Уже открытый файл
                )
                logger.debug(f"DataFrame для листа '{sheet_name}' загружен. Shape: {df.shape}")

                # --- Преобразование DataFrame в формат, ожидаемый storage ---
                raw_data_list = []
                # Итерируемся по строкам и столбцам DataFrame
                for index, row in df.iterrows():
                    for col_name, value in row.items():
                        # Pandas использует 0-based индексацию для строк
                        # Для координаты ячейки нужно учесть start_row_pandas и header_pandas
                        # row_index_pandas - это индекс строки в DataFrame (0-based)
                        # col_index_pandas - это индекс столбца в DataFrame (0-based)
                        # excel_row_index - это номер строки в Excel (1-based)
                        # excel_col_letter - это буква столбца в Excel
                        row_index_pandas = index
                        try:
                            col_index_pandas = df.columns.get_loc(col_name)
                        except KeyError:
                            logger.warning(f"Столбец '{col_name}' не найден в columns. Пропущен.")
                            continue

                        # Исправлено: Убедимся, что col_index_pandas - это int перед +1
                        if not isinstance(col_index_pandas, int):
                            logger.warning(f"col_index_pandas для '{col_name}' не является int: {col_index_pandas}. Пропущен.")
                            continue
                        
                        # Исправлено: Убедимся, что row_index_pandas - это int перед +1
                        if not isinstance(row_index_pandas, int):
                            logger.warning(f"row_index_pandas для строки {index} не является int: {row_index_pandas}. Пропущен.")
                            continue

                        # Преобразуем индексы pandas в координаты Excel
                        # Смещение для строк: start_row_pandas + header_pandas + 1 (так как Excel 1-based)
                        offset_for_header_and_skiprows = (start_row_pandas or 0) + (1 if header_pandas is not None else 0)
                        excel_row_index = row_index_pandas + 1 + offset_for_header_and_skiprows # +1 для перехода к 1-based
                        # Преобразуем индекс столбца в букву
                        # Исправлено: Используем импортированную функцию get_column_letter
                        excel_col_letter = get_column_letter(col_index_pandas + 1 + (start_col_pandas or 0)) # +1 для перехода к 1-based, + start_col_pandas
                        cell_address = f"{excel_col_letter}{excel_row_index}"

                        # Обработка значений NaN/null
                        if pd.isna(value):
                            # В БД можно сохранить как None или специальную строку
                            processed_value = None
                        else:
                            processed_value = value

                        raw_data_list.append({
                            "cell_address": cell_address,
                            "value": processed_value,
                        })

                # --- Сохранение "сырых данных" ---
                if not app_controller.storage.save_sheet_raw_data(sheet_name, raw_data_list):
                    logger.error(f"Не удалось сохранить 'сырые данные' для листа '{sheet_name}'.")
                    # В зависимости от требований, можно вернуть False или продолжить с другими листами
                    # Пока продолжим
                else:
                    logger.info(f"Быстрые 'сырые данные' для листа '{sheet_name}' успешно импортированы.")

        logger.info(f"БЫСТРЫЙ импорт 'сырых' данных из '{file_path}' завершён.")
        return True

    except ImportError:
        logger.error("Библиотека pandas не установлена. Быстрый импорт невозможен.")
        return False
    except Exception as e:
        logger.error(f"Ошибка при БЫСТРОМ импорте 'сырых' данных из файла '{file_path}': {e}", exc_info=True)
        return False
