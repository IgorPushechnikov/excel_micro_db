# src/exporter/xlsx_exporter.py
"""
Модуль для экспорта данных проекта Excel Micro DB в новый Excel-файл с использованием XlsxWriter.
"""

import xlsxwriter
import logging
from typing import Dict, Any, List, Optional
from pathlib import Path

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
import sys
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

def export_project_to_excel_xlsxwriter(project_data: Dict[str, Any], output_file_path: str) -> bool:
    """
    Экспортирует данные проекта в новый Excel-файл с использованием XlsxWriter.
    
    Args:
        project_data (Dict[str, Any]): Данные проекта, загруженные из БД.
        output_file_path (str): Путь к файлу Excel, который будет создан.
        
    Returns:
        bool: True, если экспорт прошёл успешно, иначе False.
    """
    try:
        # Создаем новый Excel-файл
        workbook = xlsxwriter.Workbook(output_file_path)
        
        # Получаем информацию о проекте
        project_info = project_data.get("project_info", {})
        project_name = project_info.get("name", "Unknown Project")
        
        logger.info(f"Начало экспорта проекта '{project_name}' в файл '{output_file_path}' с использованием XlsxWriter.")
        
        # Получаем данные листов
        sheets_data = project_data.get("sheets", {})
        
        # Создаем листы в новом файле
        for sheet_name, sheet_info in sheets_data.items():
            logger.debug(f"Экспорт листа: {sheet_name}")
            
            # Создаем лист
            worksheet = workbook.add_worksheet(sheet_name)
            
            # Экспортируем структуру и данные
            _export_sheet_data(worksheet, sheet_info)
            
            # Экспортируем формулы
            _export_sheet_formulas(worksheet, sheet_info)
            
            # Экспортируем стили
            _export_sheet_styles(workbook, worksheet, sheet_info)
            
            # Экспортируем объединенные ячейки
            _export_sheet_merged_cells(worksheet, sheet_info)
            
            # TODO: Экспортируем диаграммы (при необходимости)
            # _export_sheet_charts(workbook, worksheet, sheet_info)
            
        # Закрываем файл
        workbook.close()
        logger.info(f"Экспорт проекта '{project_name}' завершен успешно. Файл сохранен как '{output_file_path}'.")
        return True
        
    except Exception as e:
        logger.error(f"Ошибка при экспорте проекта в файл '{output_file_path}': {e}", exc_info=True)
        return False

def _export_sheet_data(worksheet, sheet_info: Dict[str, Any]) -> None:
    """Экспортирует структуру и данные листа."""
    try:
        # Получаем редактируемые данные
        editable_data = sheet_info.get("editable_data", {})
        column_names = editable_data.get("column_names", [])
        rows = editable_data.get("rows", [])
        
        if not column_names:
            logger.warning(f"Нет данных для экспорта на листе.")
            return
            
        # Записываем заголовки
        for col_idx, col_name in enumerate(column_names):
            worksheet.write(0, col_idx, col_name)
            
        # Записываем данные
        for row_idx, row_data in enumerate(rows, start=1):
            # row_data - это словарь {имя_колонки: значение}
            for col_idx, col_name in enumerate(column_names):
                value = row_data.get(col_name, "")
                worksheet.write(row_idx, col_idx, value)
                
    except Exception as e:
        logger.error(f"Ошибка при экспорте данных листа: {e}")

def _export_sheet_formulas(worksheet, sheet_info: Dict[str, Any]) -> None:
    """Экспортирует формулы листа."""
    try:
        formulas_data = sheet_info.get("formulas", [])
        
        for formula_info in formulas_data:
            cell_address = formula_info.get("cell", "")  # Например, "F2"
            formula = formula_info.get("formula", "")    # Например, "=SUM(B2:E2)"
            
            if cell_address and formula:
                # XlsxWriter позволяет записывать формулы напрямую
                worksheet.write_formula(cell_address, formula)
                
    except Exception as e:
        logger.error(f"Ошибка при экспорте формул листа: {e}")

def _export_sheet_styles(workbook, worksheet, sheet_info: Dict[str, Any]) -> None:
    """Экспортирует стили листа."""
    try:
        # TODO: Реализовать преобразование стилей из формата БД в формат XlsxWriter
        # Это потребует детального анализа структуры styled_ranges и связанных таблиц
        styled_ranges_data = sheet_info.get("styled_ranges", [])
        
        # Пример создания формата в XlsxWriter
        # header_format = workbook.add_format({
        #     'bold': True,
        #     'text_wrap': True,
        #     'valign': 'top',
        #     'fg_color': '#D7E4BC',
        #     'border': 1
        # })
        # worksheet.set_row(0, None, header_format)
        
        # Для каждого стиля из БД нужно создать соответствующий формат XlsxWriter
        # и применить его к диапазону ячеек
        
    except Exception as e:
        logger.error(f"Ошибка при экспорте стилей листа: {e}")

def _export_sheet_merged_cells(worksheet, sheet_info: Dict[str, Any]) -> None:
    """Экспортирует объединенные ячейки листа."""
    try:
        merged_cells_data = sheet_info.get("merged_cells", [])
        
        for range_address in merged_cells_data:
            if range_address:
                # XlsxWriter использует метод merge_range для объединения ячеек
                # Нужно определить значение для объединенного диапазона
                # Пока просто объединяем без значения
                worksheet.merge_range(range_address, "")
                
    except Exception as e:
        logger.error(f"Ошибка при экспорте объединенных ячеек листа: {e}")

# Точка входа для тестирования
if __name__ == "__main__":
    # Простой тест
    print("Тестирование xlsx_exporter")