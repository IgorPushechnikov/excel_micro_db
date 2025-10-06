# backend/core/app_controller_data_import.py
"""
Модуль, содержащий функции для импорта различных типов данных из Excel-файла
в БД проекта через AppController.

Функции предназначены для вызова из AppController для реализации
импорта "по типам" (данные, стили, диаграммы, формулы) и "по режимам"
(всё, выборочно, частями).

ПРИМЕЧАНИЕ: Функции теперь принимают экземпляр ProjectDBStorage напрямую,
чтобы обеспечить потокобезопасность.
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

# Импортируем ProjectDBStorage
from backend.storage.base import ProjectDBStorage

logger = get_logger(__name__)


# --- Вспомогательные функции ---

# УДАЛЯЕМ _get_sheet_id_by_name, так как storage.save_sheet делает то же самое лучше.
# def _get_sheet_id_by_name(storage, sheet_name: str) -> Optional[int]:
#     """
#     Вспомогательная функция для получения sheet_id по имени листа.
#     """
#     if not storage or not storage.connection:
#         return None
#
#     try:
#         cursor = storage.connection.cursor()
#         # Предполагаем project_id = 1
#         cursor.execute("SELECT sheet_id FROM sheets WHERE name = ? AND project_id = 1", (sheet_name,))
#         result = cursor.fetchone()
#         return result[0] if result else None
#     except Exception as e:
#         logger.error(f"Ошибка при получении sheet_id для листа '{sheet_name}': {e}")
#         return None


# --- Функции для импорта "всё" по типам ---

def import_raw_data_from_excel(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только "сырые" данные (значения ячеек) из Excel-файла в БД проекта.

    Args:
        storage: Экземпляр ProjectDBStorage для сохранения данных.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not storage:
        logger.error("Экземпляр ProjectDBStorage не предоставлен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта 'сырых' данных из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import_orig = options.get('sheets', []) if options else []
        if not sheets_to_import_orig:
            sheets_to_import_orig = workbook.sheetnames

        # Явно приводим элементы к str для устранения ошибки Pylance
        sheets_to_import: List[str] = [str(name) for name in sheets_to_import_orig]

        for sheet_name_orig in sheets_to_import:
            # Убедимся, что имя листа - строка
            sheet_name: str = str(sheet_name_orig)
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт 'сырых' данных с листа: {sheet_name}")
            
            # --- НОВОЕ: Гарантируем, что запись о листе существует ---
            # Предполагаем project_id = 1 для MVP
            sheet_id = storage.save_sheet(project_id=1, sheet_name=sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось создать/получить ID для листа '{sheet_name}'. Пропущен.")
                return False # Возвращаем False при ошибке создания записи о листе
            # --- КОНЕЦ НОВОГО ---

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

            if not storage.save_sheet_raw_data(sheet_name, raw_data_list):
                logger.error(f"Не удалось сохранить 'сырые данные' для листа '{sheet_name}'.")
                return False  # Возвращаем False при ошибке

        logger.info(f"Импорт 'сырых' данных из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте 'сырых' данных из файла '{file_path}': {e}", exc_info=True)
        return False


def import_styles_from_excel(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только стили из Excel-файла в БД проекта.

    Args:
        storage: Экземпляр ProjectDBStorage для сохранения данных.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not storage:
        logger.error("Экземпляр ProjectDBStorage не предоставлен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта стилей из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import_orig = options.get('sheets', []) if options else []
        if not sheets_to_import_orig:
            sheets_to_import_orig = workbook.sheetnames

        # Явно приводим элементы к str для устранения ошибки Pylance
        sheets_to_import: List[str] = [str(name) for name in sheets_to_import_orig]

        import json

        for sheet_name_orig in sheets_to_import:
            # Убедимся, что имя листа - строка
            sheet_name: str = str(sheet_name_orig)
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт стилей с листа: {sheet_name}")
            
            # --- НОВОЕ: Гарантируем, что запись о листе существует ---
            # Предполагаем project_id = 1 для MVP
            sheet_id = storage.save_sheet(project_id=1, sheet_name=sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось создать/получить ID для листа '{sheet_name}'. Пропущен.")
                return False # Возвращаем False при ошибке создания записи о листе
            # --- КОНЕЦ НОВОГО ---

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

            if not storage.save_sheet_styles(sheet_id, styles_to_save):
                logger.error(f"Не удалось сохранить стили для листа '{sheet_name}' (ID: {sheet_id}).")
                return False  # Возвращаем False при ошибке

        logger.info(f"Импорт стилей из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте стилей из файла '{file_path}': {e}", exc_info=True)
        return False


def import_charts_from_excel(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только диаграммы из Excel-файла в БД проекта.

    Args:
        storage: Экземпляр ProjectDBStorage для сохранения данных.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not storage:
        logger.error("Экземпляр ProjectDBStorage не предоставлен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта диаграмм из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import_orig = options.get('sheets', []) if options else []
        if not sheets_to_import_orig:
            sheets_to_import_orig = workbook.sheetnames

        # Явно приводим элементы к str для устранения ошибки Pylance
        sheets_to_import: List[str] = [str(name) for name in sheets_to_import_orig]

        for sheet_name_orig in sheets_to_import:
            # Убедимся, что имя листа - строка
            sheet_name: str = str(sheet_name_orig)
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт диаграмм с листа: {sheet_name}")
            
            # --- НОВОЕ: Гарантируем, что запись о листе существует ---
            # Предполагаем project_id = 1 для MVP
            sheet_id = storage.save_sheet(project_id=1, sheet_name=sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось создать/получить ID для листа '{sheet_name}'. Пропущен.")
                return False # Возвращаем False при ошибке создания записи о листе
            # --- КОНЕЦ НОВОГО ---

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

            if not storage.save_sheet_charts(sheet_id, charts_list):
                logger.error(f"Не удалось сохранить диаграммы для листа '{sheet_name}' (ID: {sheet_id}).")
                return False  # Возвращаем False при ошибке

        logger.info(f"Импорт диаграмм из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте диаграмм из файла '{file_path}': {e}", exc_info=True)
        return False


def import_formulas_from_excel(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует только формулы из Excel-файла в БД проекта.

    Args:
        storage: Экземпляр ProjectDBStorage для сохранения данных.
        file_path (str): Путь к Excel-файлу для импорта.
        options (Optional[Dict[str, Any]]): Опции импорта.

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not storage:
        logger.error("Экземпляр ProjectDBStorage не предоставлен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта формул из Excel-файла: {file_path}")

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        sheets_to_import_orig = options.get('sheets', []) if options else []
        if not sheets_to_import_orig:
            sheets_to_import_orig = workbook.sheetnames

        # Явно приводим элементы к str для устранения ошибки Pylance
        sheets_to_import: List[str] = [str(name) for name in sheets_to_import_orig]

        for sheet_name_orig in sheets_to_import:
            # Убедимся, что имя листа - строка
            sheet_name: str = str(sheet_name_orig)
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт формул с листа: {sheet_name}")
            
            # --- НОВОЕ: Гарантируем, что запись о листе существует ---
            # Предполагаем project_id = 1 для MVP
            sheet_id = storage.save_sheet(project_id=1, sheet_name=sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось создать/получить ID для листа '{sheet_name}'. Пропущен.")
                return False # Возвращаем False при ошибке создания записи о листе
            # --- КОНЕЦ НОВОГО ---

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

            if not storage.save_sheet_formulas(sheet_id, formulas_list):
                logger.error(f"Не удалось сохранить формулы для листа '{sheet_name}' (ID: {sheet_id}).")
                return False  # Возвращаем False при ошибке

        logger.info(f"Импорт формул из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте формул из файла '{file_path}': {e}", exc_info=True)
        return False


# --- Функции для импорта "выборочно" по типам ---

# Заглушка для выборочного импорта. Реализация будет аналогична полному импорту,
# но с фильтрацией по листам/диапазонам, переданным в options.
# Например, options = {'sheets': ['Sheet1'], 'start_row': 1, 'end_row': 100}

def import_raw_data_from_excel_selective(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует "сырые" данные выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт сырых данных выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    # Реализация будет аналогична import_raw_data_from_excel, но с фильтрацией
    # на основе options (sheets, start_row, end_row, start_col, end_col)
    # Пока возвращаем True как успешное выполнение заглушки
    return True

def import_styles_from_excel_selective(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует стили выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт стилей выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    return True

def import_charts_from_excel_selective(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует диаграммы выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт диаграмм выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    return True

def import_formulas_from_excel_selective(storage: ProjectDBStorage, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    """
    Импортирует формулы выборочно из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт формул выборочно из '{file_path}' с опциями {options}. (Заглушка)")
    return True


# --- Функции для импорта "частями" по типам ---

# Заглушка для импорта частями. Реализация будет разбивать большой файл
# на части и вызывать соответствующую функцию импорта для каждой части.
# Например, импортировать по 1000 строк за раз.

def import_raw_data_from_excel_in_chunks(storage: ProjectDBStorage, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует "сырые" данные (значения ячеек) частями из Excel-файла в БД проекта.
    Использует openpyxl и параметр 'chunk_size_rows' из chunk_options.

    Args:
        storage: Экземпляр ProjectDBStorage для сохранения данных.
        file_path (str): Путь к Excel-файлу для импорта.
        chunk_options (Dict[str, Any]): Опции для разбиения на части.
            {
                'chunk_size_rows': int, # Количество строк в одной части (по умолчанию 1000)
                'sheets': List[str]     # Список имён листов для импорта. Если пуст, все.
            }

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not storage:
        logger.error("Экземпляр ProjectDBStorage не предоставлен. Невозможно выполнить импорт.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл для импорта не найден: {file_path}")
        return False

    logger.info(f"Начало импорта 'сырых' данных частями из Excel-файла: {file_path} с опциями {chunk_options}")

    # Установим значения по умолчанию
    chunk_size = chunk_options.get('chunk_size_rows', 1000)
    sheets_to_import_orig = chunk_options.get('sheets', [])

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False, read_only=True) # read_only=True может помочь с памятью
        logger.debug(f"Книга '{file_path}' успешно открыта в режиме 'read_only'.")

        if not sheets_to_import_orig:
            sheets_to_import_orig = workbook.sheetnames

        # Явно приводим элементы к str
        sheets_to_import: List[str] = [str(name) for name in sheets_to_import_orig]

        for sheet_name_orig in sheets_to_import:
            sheet_name: str = str(sheet_name_orig)
            if sheet_name not in workbook.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле '{file_path}'. Пропущен.")
                continue

            logger.info(f"Импорт 'сырых' данных с листа: {sheet_name} частями по {chunk_size} строк")

            # --- НОВОЕ: Гарантируем, что запись о листе существует ---
            sheet_id = storage.save_sheet(project_id=1, sheet_name=sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось создать/получить ID для листа '{sheet_name}'. Пропущен.")
                return False
            # --- КОНЕЦ НОВОГО ---

            sheet: Worksheet = workbook[sheet_name]
            total_rows = sheet.max_row
            logger.debug(f"Обнаружено {total_rows} строк на листе '{sheet_name}'.")

            start_row = 1 # openpyxl использует 1-based индексацию
            while start_row <= total_rows:
                end_row = min(start_row + chunk_size - 1, total_rows)
                logger.debug(f"Обработка строки {start_row} - {end_row}.")

                raw_data_list = []
                # Используем iter_rows с указанием min_row и max_row для "части"
                for row in sheet.iter_rows(min_row=start_row, max_row=end_row, values_only=False):
                    for cell in row:
                        if cell.value is not None or cell.data_type == 'f': # Сохраняем значения и формулы
                            data_item = {
                                "cell_address": cell.coordinate,
                                "value": cell.value,
                            }
                            raw_data_list.append(data_item)

                # Сохраняем "часть" данных в БД
                if not storage.save_sheet_raw_data(sheet_name, raw_data_list):
                    logger.error(f"Не удалось сохранить 'сырые данные' для листа '{sheet_name}' (часть строки {start_row}-{end_row}).")
                    return False

                logger.debug(f"Сохранена часть данных с {start_row} по {end_row} для листа '{sheet_name}'.")

                start_row = end_row + 1 # Переходим к следующей части

        logger.info(f"Импорт 'сырых' данных частями из '{file_path}' завершён.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте 'сырых' данных частями из файла '{file_path}': {e}", exc_info=True)
        return False

def import_styles_from_excel_in_chunks(storage: ProjectDBStorage, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует стили частями из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт стилей частями из '{file_path}' с опциями {chunk_options}. (Заглушка)")
    return True

def import_charts_from_excel_in_chunks(storage: ProjectDBStorage, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует диаграммы частями из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт диаграмм частями из '{file_path}' с опциями {chunk_options}. (Заглушка)")
    return True

def import_formulas_from_excel_in_chunks(storage: ProjectDBStorage, file_path: str, chunk_options: Dict[str, Any]) -> bool:
    """
    Импортирует формулы частями из Excel-файла в БД проекта.
    """
    logger.info(f"Импорт формул частями из '{file_path}' с опциями {chunk_options}. (Заглушка)")
    return True
