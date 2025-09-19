# src/exporter/data_and_formulas_exporter.py
"""Модуль для экспорта данных и формул листа Excel."""
import sys
from pathlib import Path
from typing import Dict, Any, List, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string

# Добавляем корень проекта в путь поиска модулей
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

def export_sheet_structure(ws: Worksheet, structure_data: List[Dict[str, Any]]) -> None:
    """Экспортирует структуру листа (заголовки столбцов)."""
    logger.debug(f"[ЭКСПОРТ_СТРУКТУРЫ] Начало экспорта структуры для листа '{ws.title}'")
    logger.debug(f"[ЭКСПОРТ_СТРУКТУРЫ] Получено {len(structure_data)} элементов структуры.")
    
    if not structure_data:
        logger.debug("[ЭКСПОРТ_СТРУКТУРЫ] Нет данных структуры для экспорта.")
        return

    logger.debug("[ЭКСПОРТ_СТРУКТУРЫ] Начало итерации по данным структуры...")
    for i, col_info in enumerate(structure_data, start=1):
        logger.debug(f"[ЭКСПОРТ_СТРУКТУРЫ] Обработка элемента структуры {i}: {col_info}")
        col_index = col_info.get("column_index", i)
        col_name = col_info.get("column_name", f"Column_{col_index}")
        
        # Получаем букву столбца
        try:
            col_letter = get_column_letter(col_index)
        except Exception as e:
            logger.warning(f"[ЭКСПОРТ_СТРУКТУРЫ] Ошибка получения буквы столбца для индекса {col_index}: {e}. Используется 'A'.")
            col_letter = 'A'
            
        cell_address = f"{col_letter}1" # Заголовки в первой строке
        logger.debug(f"[ЭКСПОРТ_СТРУКТУРЫ] Запись заголовка '{col_name}' в ячейку {cell_address}")
        ws[cell_address] = col_name
        
    logger.debug(f"[ЭКСПОРТ_СТРУКТУРЫ] === Конец экспорта структуры для листа '{ws.title}' ===")

def export_sheet_raw_data(ws: Worksheet, raw_data_info: Dict[str, Any]) -> None:
    """Экспортирует сырые данные листа."""
    logger.debug(f"[ЭКСПОРТ_ДАННЫХ] Начало экспорта сырых данных для листа '{ws.title}'")
    
    column_names = raw_data_info.get("column_names", [])
    rows_data = raw_data_info.get("rows", [])
    
    logger.debug(f"[ЭКСПОРТ_ДАННЫХ] Получено {len(column_names)} имен столбцов и {len(rows_data)} строк данных.")
    
    if not column_names:
        logger.debug("[ЭКСПОРТ_ДАННЫХ] Нет имен столбцов для экспорта.")
        return

    # Записываем данные строк, начиная со строки 2 (так как строка 1 - для структуры/заголовков)
    for row_idx, row_dict in enumerate(rows_data, start=2): 
        logger.debug(f"[ЭКСПОРТ_ДАННЫХ] Обработка строки данных {row_idx - 1} (Excel строка {row_idx}): {list(row_dict.keys())}")
        for col_idx, col_name in enumerate(column_names, start=1):
            col_letter = get_column_letter(col_idx)
            cell_address = f"{col_letter}{row_idx}"
            cell_value = row_dict.get(col_name, "")
            logger.debug(f"[ЭКСПОРТ_ДАННЫХ] Запись значения '{cell_value}' в ячейку {cell_address}")
            ws[cell_address] = cell_value
            
    logger.debug(f"[ЭКСПОРТ_ДАННЫХ] === Конец экспорта сырых данных для листа '{ws.title}' ===")

def export_sheet_formulas(ws: Worksheet, formulas_data: List[Dict[str, Any]]) -> None:
    """Экспортирует формулы листа."""
    logger.debug(f"[ЭКСПОРТ_ФОРМУЛ] Начало экспорта формул для листа '{ws.title}'")
    logger.debug(f"[ЭКСПОРТ_ФОРМУЛ] Получено {len(formulas_data)} формул.")
    
    if not formulas_data:
        logger.debug("[ЭКСПОРТ_ФОРМУЛ] Нет формул для экспорта.")
        return
        
    logger.debug("[ЭКСПОРТ_ФОРМУЛ] Начало итерации по формулам...")
    for i, formula_info in enumerate(formulas_data, start=1):
        logger.debug(f"[ЭКСПОРТ_ФОРМУЛ] Обработка формулы {i}: {formula_info}")
        cell_address = formula_info.get("cell")
        formula_text = formula_info.get("formula")
        
        if not cell_address or not formula_text:
            logger.warning(f"[ЭКСПОРТ_ФОРМУЛ] Пропущена формула {i} из-за отсутствия адреса или текста: {formula_info}")
            continue
            
        logger.debug(f"[ЭКСПОРТ_ФОРМУЛ] Попытка записи формулы в {cell_address}: {formula_text}")
        try:
            # Убедимся, что формула начинается со знака =
            if not formula_text.startswith("="):
                formula_text = f"={formula_text}"
            ws[cell_address] = formula_text
            logger.debug(f"[ЭКСПОРТ_ФОРМУЛ] Успешно записана формула в {cell_address}")
        except Exception as e:
            logger.error(f"[ЭКСПОРТ_ФОРМУЛ] Ошибка записи формулы в {cell_address}: {e}")
            
    logger.debug(f"[ЭКСПОРТ_ФОРМУЛ] === Конец экспорта формул для листа '{ws.title}' ===")
