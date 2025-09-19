# src/analyzer/logic_documentation.py
"""Анализатор логики Excel файлов для Excel Micro DB.
Извлекает структуру данных, формулы, зависимости, стили и диаграммы из Excel файлов.
"""
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
# === ИСПРАВЛЕНО: Импорт coordinate_from_string из правильного модуля ===
from openpyxl.utils.cell import coordinate_from_string
# ========================================================================
# === ИМПОРТЫ ДЛЯ РАБОТЫ СО СТИЛЯМИ ===
from openpyxl.styles import Font, Fill, Border, PatternFill, Side, Alignment, Protection
# === КОНЕЦ ИМПОРТОВ ДЛЯ СТИЛЕЙ ===
# === ИМПОРТЫ ДЛЯ СОЗДАНИЯ ДИАГРАММ С ДАННЫМИ ===
from openpyxl.chart import BarChart, PieChart, LineChart, ScatterChart, AreaChart, PieChart3D
from openpyxl.chart.title import Title
from openpyxl.chart.layout import Layout
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, RegularTextRun
# === КОНЕЦ ИМПОРТОВ ===
from openpyxl.chart.series import Series
import yaml
from datetime import datetime
import re
import logging

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

# Получаем логгер для этого модуля
logger = get_logger(__name__)

# - ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ -

def load_documentation_template(template_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Загружает шаблон документации из YAML файла.
    Если путь не указан, использует путь по умолчанию.
    Args:
        template_path (str, optional): Путь к файлу шаблона. Defaults to None.
    Returns:
        Dict[str, Any]: Загруженный шаблон документации.
    """
    if template_path is None:
        # Определяем путь к шаблону по умолчанию относительно этого файла
        template_path = project_root / "templates" / "documentation_template.yaml"
        
    logger.debug(f"Загрузка шаблона документации из: {template_path}")
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            template = yaml.safe_load(f)
        logger.debug("Шаблон документации успешно загружен")
        return template
    except FileNotFoundError:
        logger.error(f"Файл шаблона документации не найден: {template_path}")
        raise
    except yaml.YAMLError as e:
        logger.error(f"Ошибка форматирования YAML в файле шаблона: {e}")
        raise

def get_cell_address(row: int, col: int) -> str:
    """
    Преобразует числовые координаты ячейки в адрес Excel (например, 1,1 -> A1).
    Args:
        row (int): Номер строки (начиная с 1)
        col (int): Номер столбца (начиная с 1)
    Returns:
        str: Адрес ячейки в формате Excel (например, 'A1')
    """
    try:
        column_letter = get_column_letter(col)
        return f"{column_letter}{row}"
    except Exception as e:
        logger.error(f"Ошибка при преобразовании координат ({row}, {col}) в адрес ячейки: {e}")
        return f"ERR_{row}_{col}"

# === НОВОЕ: Функция для извлечения текста из объекта Title openpyxl ===
def _extract_title_text(title_obj: Optional[Title]) -> str:
    """
    Извлекает текст из объекта Title openpyxl.
    Args:
        title_obj (openpyxl.chart.title.Title, optional): Объект заголовка.
    Returns:
        str: Извлеченный текст или пустая строка.
    """
    if not title_obj:
        return ""
    try:
        # Попытка получить текст из tx.rich.paragraphs (наиболее распространенный случай)
        if hasattr(title_obj, 'tx') and title_obj.tx:
            if hasattr(title_obj.tx, 'rich') and title_obj.tx.rich:
                text_parts = []
                for paragraph in title_obj.tx.rich.paragraphs:
                    if hasattr(paragraph, 'runs'):
                        for run in paragraph.runs:
                            if hasattr(run, 'text'):
                                text_parts.append(run.text)
                return "".join(text_parts).strip()
            # Если rich нет, попробуем получить текст напрямую из tx (например, strRef)
            # elif hasattr(title_obj.tx, 'strRef') and title_obj.tx.strRef:
            #     # strRef.f содержит формулу, strRef.strCache содержит кэшированные значения
            #     # Для простоты можно вернуть формулу или первое значение из кэша
            #     if hasattr(title_obj.tx.strRef, 'strCache') and title_obj.tx.strRef.strCache:
            #         pt_list = getattr(title_obj.tx.strRef.strCache, 'pt', [])
            #         if pt_list:
            #             first_pt = pt_list[0]
            #             if hasattr(first_pt, 'v'):
            #                 return str(first_pt.v).strip()
    except Exception as e:
        logger.warning(f"Ошибка при извлечении текста из Title: {e}")
    # Если стандартные методы не сработали, попробуем str()
    try:
        return str(title_obj).strip()
    except:
        pass
    return ""

# === НОВОЕ: Функция для извлечения атрибутов стиля ячейки ===
def _extract_style_attributes(cell) -> Dict[str, Any]:
    """
    Извлекает атрибуты стиля ячейки в структурированном виде.
    Args:
        cell (openpyxl.cell.Cell): Объект ячейки.
    Returns:
        Dict[str, Any]: Словарь с атрибутами стиля.
    """
    style_attrs = {
        # Font
        "font_name": None, "font_sz": None, "font_b": None, "font_i": None,
        "font_u": None, "font_strike": None, "font_color": None,
        "font_color_theme": None, "font_color_tint": None,
        "font_vert_align": None, "font_scheme": None,
        # PatternFill
        "fill_pattern_type": None, "fill_fg_color": None,
        "fill_fg_color_theme": None, "fill_fg_color_tint": None,
        "fill_bg_color": None, "fill_bg_color_theme": None, "fill_bg_color_tint": None,
        # Border (основные стороны)
        "border_left_style": None, "border_left_color": None,
        "border_right_style": None, "border_right_color": None,
        "border_top_style": None, "border_top_color": None,
        "border_bottom_style": None, "border_bottom_color": None,
        # Alignment
        "alignment_horizontal": None, "alignment_vertical": None,
        "alignment_text_rotation": None, "alignment_wrap_text": None,
        "alignment_shrink_to_fit": None, "alignment_indent": None,
        # Protection
        "protection_locked": None, "protection_hidden": None,
    }

    try:
        if cell.font:
            style_attrs["font_name"] = getattr(cell.font, 'name', None)
            style_attrs["font_sz"] = float(getattr(cell.font, 'sz', 0)) if getattr(cell.font, 'sz', None) is not None else None
            style_attrs["font_b"] = int(bool(getattr(cell.font, 'b', False)))
            style_attrs["font_i"] = int(bool(getattr(cell.font, 'i', False)))
            style_attrs["font_u"] = getattr(cell.font, 'u', None)
            style_attrs["font_strike"] = int(bool(getattr(cell.font, 'strike', False)))
            if getattr(cell.font, 'color', None):
                if getattr(cell.font.color, 'type', None) == 'rgb':
                    style_attrs["font_color"] = getattr(cell.font.color, 'rgb', None)
                elif getattr(cell.font.color, 'type', None) == 'theme':
                    style_attrs["font_color_theme"] = int(getattr(cell.font.color, 'theme', 0)) if getattr(cell.font.color, 'theme', None) is not None else None
                    style_attrs["font_color_tint"] = float(getattr(cell.font.color, 'tint', 0.0)) if getattr(cell.font.color, 'tint', None) is not None else None
            style_attrs["font_vert_align"] = getattr(cell.font, 'vertAlign', None)
            style_attrs["font_scheme"] = getattr(cell.font, 'scheme', None)

        if cell.fill:
            if isinstance(cell.fill, PatternFill):
                style_attrs["fill_pattern_type"] = getattr(cell.fill, 'patternType', None)
                if getattr(cell.fill, 'fgColor', None):
                    if getattr(cell.fill.fgColor, 'type', None) == 'rgb':
                        style_attrs["fill_fg_color"] = getattr(cell.fill.fgColor, 'rgb', None)
                    elif getattr(cell.fill.fgColor, 'type', None) == 'theme':
                        style_attrs["fill_fg_color_theme"] = int(getattr(cell.fill.fgColor, 'theme', 0)) if getattr(cell.fill.fgColor, 'theme', None) is not None else None
                        style_attrs["fill_fg_color_tint"] = float(getattr(cell.fill.fgColor, 'tint', 0.0)) if getattr(cell.fill.fgColor, 'tint', None) is not None else None
                if getattr(cell.fill, 'bgColor', None):
                    if getattr(cell.fill.bgColor, 'type', None) == 'rgb':
                        style_attrs["fill_bg_color"] = getattr(cell.fill.bgColor, 'rgb', None)
                    elif getattr(cell.fill.bgColor, 'type', None) == 'theme':
                        style_attrs["fill_bg_color_theme"] = int(getattr(cell.fill.bgColor, 'theme', 0)) if getattr(cell.fill.bgColor, 'theme', None) is not None else None
                        style_attrs["fill_bg_color_tint"] = float(getattr(cell.fill.bgColor, 'tint', 0.0)) if getattr(cell.fill.bgColor, 'tint', None) is not None else None

        if cell.border:
             # Извлекаем основные стороны
             for side_name in ['left', 'right', 'top', 'bottom']:
                 side_obj = getattr(cell.border, side_name, None)
                 if side_obj and hasattr(side_obj, 'style'):
                     style_attrs[f"border_{side_name}_style"] = getattr(side_obj, 'style', None)
                     if getattr(side_obj, 'color', None):
                         if getattr(side_obj.color, 'type', None) == 'rgb':
                             style_attrs[f"border_{side_name}_color"] = getattr(side_obj.color, 'rgb', None)

        if cell.alignment:
            style_attrs["alignment_horizontal"] = getattr(cell.alignment, 'horizontal', None)
            style_attrs["alignment_vertical"] = getattr(cell.alignment, 'vertical', None)
            style_attrs["alignment_text_rotation"] = int(getattr(cell.alignment, 'textRotation', 0)) if getattr(cell.alignment, 'textRotation', None) is not None else None
            style_attrs["alignment_wrap_text"] = int(bool(getattr(cell.alignment, 'wrapText', False)))
            style_attrs["alignment_shrink_to_fit"] = int(bool(getattr(cell.alignment, 'shrinkToFit', False)))
            style_attrs["alignment_indent"] = int(getattr(cell.alignment, 'indent', 0)) if getattr(cell.alignment, 'indent', None) is not None else None

        if cell.protection:
            style_attrs["protection_locked"] = int(bool(getattr(cell.protection, 'locked', True))) # locked по умолчанию True
            style_attrs["protection_hidden"] = int(bool(getattr(cell.protection, 'hidden', False)))

    except Exception as e:
        logger.error(f"Ошибка при извлечении атрибутов стиля для ячейки {cell.coordinate}: {e}")

    # Убираем ключи со значениями None для компактности
    return {k: v for k, v in style_attrs.items() if v is not None}


# - ОСНОВНЫЕ ФУНКЦИИ АНАЛИЗА -

def analyze_sheet_structure(sheet) -> List[Dict[str, Any]]:
    """
    Анализирует структуру листа Excel: извлекает заголовки столбцов.
    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): Объект листа Excel.
    Returns:
        List[Dict[str, Any]]: Список словарей с информацией о каждом столбце.
    """
    logger.debug(f"Начало анализа структуры листа '{sheet.title}'")
    structure_info = []
    try:
        header_row = 1
        max_col = sheet.max_column
        logger.debug(f"Анализируется строка заголовков {header_row}, макс. столбец: {max_col}")

        for col in range(1, max_col + 1):
            cell = sheet.cell(row=header_row, column=col)
            column_name = str(cell.value) if cell.value is not None else f"Столбец_{col}"
            
            sample_values = []
            for sample_row in range(header_row + 1, min(header_row + 4, sheet.max_row + 1)):
                sample_cell = sheet.cell(row=sample_row, column=col)
                sample_values.append(str(sample_cell.value) if sample_cell.value is not None else None)
            
            column_info = {
                "column_name": column_name,
                "column_index": col,
                "data_type": "unknown",
                "sample_values": sample_values,
                "unique_count": 0,
                "null_count": 0,
                "description": ""
            }
            structure_info.append(column_info)
        
        logger.debug(f"Структура листа '{sheet.title}' проанализирована. Найдено {len(structure_info)} столбцов.")
        return structure_info
    except Exception as e:
        logger.error(f"Ошибка при анализе структуры листа '{sheet.title}': {e}", exc_info=True)
        return []

def analyze_sheet_raw_data(sheet) -> Dict[str, Any]:
    """
    Извлекает "сырые" данные листа, начиная со второй строки.
    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): Объект листа Excel.
    Returns:
        Dict[str, Any]: Словарь с ключами 'column_names' (список имен) и 'rows' (список словарей {'col_name': value}).
    """
    logger.debug(f"[СЫРЫЕ ДАННЫЕ] Начало извлечения сырых данных для листа '{sheet.title}'")
    raw_data_info = {"column_names": [], "rows": []}
    try:
        if sheet.max_row < 2:
            logger.debug(f"[СЫРЫЕ ДАННЫЕ] Лист '{sheet.title}' пуст или содержит только заголовки.")
            return raw_data_info

        header_row = 1
        column_names = []
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=header_row, column=col)
            col_name = str(cell.value) if cell.value is not None else f"Столбец_{col}"
            column_names.append(col_name)
        
        raw_data_info["column_names"] = column_names
        logger.debug(f"[СЫРЫЕ ДАННЫЕ] Имена столбцов: {column_names}")

        data_rows = []
        for row_num in range(2, sheet.max_row + 1):
            row_data = {}
            is_row_empty = True
            for col_idx, col_name in enumerate(column_names, start=1):
                cell = sheet.cell(row=row_num, column=col_idx)
                cell_value = cell.value
                
                if isinstance(cell_value, datetime):
                    processed_value = cell_value.isoformat()
                elif pd.isna(cell_value):
                    processed_value = None
                else:
                    processed_value = str(cell_value)
                
                if processed_value is not None:
                    is_row_empty = False
                
                row_data[col_name] = processed_value
            
            if not is_row_empty:
                data_rows.append(row_data)
        
        raw_data_info["rows"] = data_rows
        logger.debug(f"[СЫРЫЕ ДАННЫЕ] Извлечено {len(data_rows)} строк данных для листа '{sheet.title}'.")
        return raw_data_info

    except Exception as e:
        logger.error(f"[СЫРЫЕ ДАННЫЕ] Ошибка при извлечении сырых данных для листа '{sheet.title}': {e}", exc_info=True)
        return {"column_names": [], "rows": []}

# === НОВОЕ: Функция для анализа стилей ===
def analyze_sheet_styles(sheet) -> List[Dict[str, Any]]:
    """
    Анализирует стили ячеек на листе и группирует их по уникальным определениям.
    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): Объект листа Excel.
    Returns:
        List[Dict[str, Any]]: Список словарей, где каждый словарь содержит
                              'style_attributes' (dict) и 'range_address' (str).
    """
    logger.debug(f"[СТИЛИ] Начало анализа стилей для листа '{sheet.title}'")
    styled_ranges = []
    try:
        if sheet.max_row < 1 or sheet.max_column < 1:
            logger.debug(f"[СТИЛИ] Лист '{sheet.title}' пуст.")
            return styled_ranges

        # Карта для хранения стилей и соответствующих им ячеек
        # Для простоты каждая ячейка с уникальным стилем -> отдельная запись.
        # В реальном проекте стоит группировать смежные ячейки с одинаковым стилем.
        
        for row in sheet.iter_rows():
            for cell in row:
                # Получаем структурированные атрибуты стиля
                style_attrs = _extract_style_attributes(cell)
                
                # Если стиль не пустой
                if style_attrs:
                    cell_address = cell.coordinate
                    styled_ranges.append({
                        "style_attributes": style_attrs,
                        "range_address": cell_address # Можно улучшить до диапазонов A1:A1
                    })
        
        logger.debug(f"[СТИЛИ] Анализ стилей для листа '{sheet.title}' завершен. Найдено {len(styled_ranges)} записей о стилях.")
        return styled_ranges

    except Exception as e:
        logger.error(f"[СТИЛИ] Ошибка при анализе стилей для листа '{sheet.title}': {e}", exc_info=True)
        return []
# === КОНЕЦ НОВОГО ===

def parse_formula_references(formula: str, current_sheet_name: str) -> List[Dict[str, Any]]:
    """
    Парсит формулу и извлекает ссылки на ячейки/диапазоны.
    Args:
        formula (str): Строка формулы Excel.
        current_sheet_name (str): Имя текущего листа.
    Returns:
        List[Dict[str, Any]]: Список словарей с информацией о ссылках.
    """
    references = []
    if not formula or not isinstance(formula, str) or not formula.startswith('='):
        return references

    ref_pattern = re.compile(
        r"(?<!\w)"
        r"(?:"
        r"'?([^'!\s]+?)'?"
        r"!"
        r")?"
        r"(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?"
        r"(?!\w)"
    )

    matches = ref_pattern.finditer(formula)
    for match in matches:
        sheet_part = match.group(1)
        ref_start = match.group(2)
        ref_end = match.group(3)
        
        if sheet_part:
            ref_sheet_name = sheet_part.strip("'")
        else:
            ref_sheet_name = current_sheet_name

        if ref_end:
            reference_type = "range"
            reference_address = f"{ref_start}:{ref_end}"
        else:
            reference_type = "cell"
            reference_address = ref_start

        ref_info = {
            "sheet": ref_sheet_name,
            "type": reference_type,
            "address": reference_address
        }
        references.append(ref_info)
        
    return references

def analyze_sheet_formulas(sheet) -> List[Dict[str, Any]]:
    """
    Анализирует формулы на листе.
    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): Объект листа Excel.
    Returns:
        List[Dict[str, Any]]: Список словарей с информацией о формулах.
    """
    logger.debug(f"Начало анализа формул на листе '{sheet.title}'")
    formulas_info = []
    try:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value:
                    formula_string = cell.value
                    cell_address = get_cell_address(cell.row, cell.column)
                    
                    references = parse_formula_references(formula_string, sheet.title)
                    
                    formula_info = {
                        "cell": cell_address,
                        "formula": formula_string,
                        "references": references
                    }
                    formulas_info.append(formula_info)
        logger.debug(f"Анализ формул на листе '{sheet.title}' завершен. Найдено {len(formulas_info)} формул.")
        return formulas_info
    except Exception as e:
        logger.error(f"Ошибка при анализе формул на листе '{sheet.title}': {e}", exc_info=True)
        return []

def analyze_cross_sheet_references(formulas_info: List[Dict[str, Any]], current_sheet_name: str) -> List[Dict[str, Any]]:
    """
    Анализирует формулы для выявления межлистовых ссылок.
    Args:
        formulas_info (List[Dict[str, Any]]): Список информации о формулах.
        current_sheet_name (str): Имя текущего листа.
    Returns:
        List[Dict[str, Any]]: Список словарей с информацией о межлистовых ссылках.
    """
    logger.debug(f"Начало анализа межлистовых ссылок для листа '{current_sheet_name}'")
    cross_sheet_refs = []
    try:
        for formula_info in formulas_info:
            formula_cell = formula_info["cell"]
            formula_string = formula_info["formula"]
            references = formula_info.get("references", [])
            
            for ref in references:
                ref_sheet = ref["sheet"]
                ref_type = ref["type"]
                ref_address = ref["address"]
                
                if ref_sheet and ref_sheet != current_sheet_name:
                    cross_ref_info = {
                        "from_sheet": current_sheet_name,
                        "from_cell": formula_cell,
                        "from_formula": formula_string,
                        "to_sheet": ref_sheet,
                        "reference_type": ref_type,
                        "reference_address": ref_address
                    }
                    cross_sheet_refs.append(cross_ref_info)
        logger.debug(f"Анализ межлистовых ссылок для листа '{current_sheet_name}' завершен. Найдено {len(cross_sheet_refs)} ссылок.")
        return cross_sheet_refs
    except Exception as e:
        logger.error(f"Ошибка при анализе межлистовых ссылок для листа '{current_sheet_name}': {e}", exc_info=True)
        return []

# === ИЗМЕНЕНО: Улучшенная функция извлечения данных диаграммы ===
def extract_chart_data(chart, sheet) -> Dict[str, Any]:
    """
    Извлекает информацию о диаграмме в структурированном виде.
    Args:
        chart (openpyxl.chart.*): Объект диаграммы openpyxl.
        sheet (openpyxl.worksheet.worksheet.Worksheet): Лист, на котором находится диаграмма.
    Returns:
        Dict[str, Any]: Словарь с информацией о диаграмме.
    """
    logger.debug(f"[ДИАГРАММЫ] Начало извлечения данных диаграммы типа {type(chart).__name__}")
    chart_data = {
        "type": type(chart).__name__,
        "title": "",
        # Chart attributes
        "top_left_cell": "",
        "width": None,
        "height": None,
        "style": None,
        "legend_position": None,
        "auto_scaling": None,
        "plot_vis_only": None,
        # Axes
        "axes": [],
        # Series
        "series": [],
        # Data Sources (will be populated from series)
        "data_sources": []
    }
    try:
        # 1. Извлечение заголовка
        chart_data["title"] = _extract_title_text(getattr(chart, 'title', None))

        # 2. Извлечение позиции и размеров (привязки)
        anchor = chart.anchor
        if anchor:
            logger.debug(f"[ДИАГРАММЫ] Тип anchor: {type(anchor)}")
            if hasattr(anchor, '_from'):
                from_marker = anchor._from
                if from_marker:
                    try:
                        top_left_col_letter = get_column_letter(from_marker.col + 1) # openpyxl использует 0-based
                        top_left_row = from_marker.row + 1 # openpyxl использует 0-based
                        chart_data["top_left_cell"] = f"{top_left_col_letter}{top_left_row}"
                    except Exception as e:
                        logger.warning(f"[ДИАГРАММЫ] Ошибка при определении top_left_cell: {e}")
            
            # Извлечение размеров (если доступны)
            if hasattr(anchor, 'ext'):
                 ext = anchor.ext
                 if ext and hasattr(ext, 'cx') and hasattr(ext, 'cy'):
                     # cx, cy в EMU (English Metric Units). 1 inch = 914400 EMU
                     # Примерный перевод в см или пункты может потребоваться, но для БД храним как есть (REAL)
                     chart_data["width"] = float(ext.cx) if ext.cx is not None else None
                     chart_data["height"] = float(ext.cy) if ext.cy is not None else None

        # 3. Извлечение атрибутов Chart
        chart_data["style"] = int(getattr(chart, 'style', 2)) if getattr(chart, 'style', None) is not None else None
        if getattr(chart, 'legend', None) and getattr(chart.legend, 'position', None):
             chart_data["legend_position"] = str(getattr(chart.legend, 'position', ''))
        chart_data["auto_scaling"] = int(bool(getattr(chart, 'auto_scaling', False)))
        chart_data["plot_vis_only"] = int(bool(getattr(chart, 'plotVisOnly', True))) # По умолчанию True
        # dispBlanksAs и showHiddenData можно добавить аналогично

        # 4. Извлечение осей
        # chart.x_axis, chart.y_axis, chart.z_axis
        for axis_attr_name in ['x_axis', 'y_axis', 'z_axis']:
             axis_obj = getattr(chart, axis_attr_name, None)
             if axis_obj:
                 axis_info = {
                     "axis_type": axis_attr_name,
                     "ax_id": int(getattr(axis_obj, 'axId', 0)) if getattr(axis_obj, 'axId', None) is not None else None,
                     "ax_pos": str(getattr(axis_obj, 'axPos', '')),
                     "delete": int(bool(getattr(axis_obj, 'delete', False))),
                     "title": _extract_title_text(getattr(axis_obj, 'title', None)),
                     # Scaling
                     "min": None, "max": None, "orientation": None, "major_unit": None, "minor_unit": None, "log_base": None,
                     # Ticks and Labels
                     "major_tick_mark": str(getattr(axis_obj, 'majorTickMark', '')),
                     "minor_tick_mark": str(getattr(axis_obj, 'minorTickMark', '')),
                     "tick_lbl_pos": str(getattr(axis_obj, 'tickLblPos', '')),
                     # Number Format
                     "num_fmt": str(getattr(getattr(axis_obj, 'numFmt', None), 'formatCode', '')) if getattr(axis_obj, 'numFmt', None) else '',
                     # Crosses
                     "crosses": str(getattr(axis_obj, 'crosses', '')),
                     "crosses_at": float(getattr(axis_obj, 'crossesAt', 0.0)) if getattr(axis_obj, 'crossesAt', None) is not None else None,
                     # Gridlines
                     "major_gridlines": int(bool(getattr(axis_obj, 'majorGridlines', None))),
                     "minor_gridlines": int(bool(getattr(axis_obj, 'minorGridlines', None))),
                 }
                 # Извлечение scaling
                 scaling = getattr(axis_obj, 'scaling', None)
                 if scaling:
                     axis_info["min"] = float(getattr(scaling, 'min', 0.0)) if getattr(scaling, 'min', None) is not None else None
                     axis_info["max"] = float(getattr(scaling, 'max', 0.0)) if getattr(scaling, 'max', None) is not None else None
                     axis_info["orientation"] = str(getattr(scaling, 'orientation', 'minMax'))
                     axis_info["major_unit"] = float(getattr(scaling, 'majorUnit', 0.0)) if getattr(scaling, 'majorUnit', None) is not None else None
                     axis_info["minor_unit"] = float(getattr(scaling, 'minorUnit', 0.0)) if getattr(scaling, 'minorUnit', None) is not None else None
                     axis_info["log_base"] = float(getattr(scaling, 'logBase', 0.0)) if getattr(scaling, 'logBase', None) is not None else None
                 
                 chart_data["axes"].append(axis_info)

        # 5. Извлечение серий данных
        for i, series in enumerate(getattr(chart, 'series', [])):
            logger.debug(f"[ДИАГРАММЫ] Обработка серии {i+1}")
            series_info = {
                "idx": int(getattr(series, 'idx', i)) if getattr(series, 'idx', None) is not None else i,
                "order": int(getattr(series, 'order', i)) if getattr(series, 'order', None) is not None else i,
                "tx": _extract_title_text(getattr(series, 'tx', None)),
                "shape": str(getattr(series, 'shape', '')),
                "smooth": int(bool(getattr(series, 'smooth', False))),
                "invert_if_negative": int(bool(getattr(series, 'invertIfNegative', False))),
                # Data Sources for this series
                "data_points": [] # Можно добавить при необходимости
            }
            
            # Извлечение источников данных для этой серии
            # Values (Y)
            if hasattr(series, 'val') and series.val:
                 if hasattr(series.val, 'numRef') and series.val.numRef and hasattr(series.val.numRef, 'f'):
                    values_formula = series.val.numRef.f
                    chart_data["data_sources"].append({
                        "series_index": i,
                        "data_type": "values",
                        "formula": values_formula
                    })
            
            # Categories (X)
            if hasattr(series, 'cat') and series.cat:
                if hasattr(series.cat, 'strRef') and series.cat.strRef and hasattr(series.cat.strRef, 'f'):
                    categories_formula = series.cat.strRef.f
                    chart_data["data_sources"].append({
                        "series_index": i,
                        "data_type": "categories",
                        "formula": categories_formula
                    })
                elif hasattr(series.cat, 'numRef') and series.cat.numRef and hasattr(series.cat.numRef, 'f'):
                     categories_formula = series.cat.numRef.f
                     chart_data["data_sources"].append({
                         "series_index": i,
                         "data_type": "categories",
                         "formula": categories_formula
                     })

            chart_data["series"].append(series_info)
        
        logger.debug(f"[ДИАГРАММЫ] Извлечено {len(chart_data['series'])} серий и {len(chart_data['data_sources'])} источников данных для диаграммы.")
        logger.debug(f"[ДИАГРАММЫ] Завершено извлечение данных диаграммы.")
        return chart_data

    except Exception as e:
        logger.error(f"[ДИАГРАММЫ] Ошибка при извлечении данных диаграммы: {e}", exc_info=True)
        return chart_data 
# === КОНЕЦ ИЗМЕНЕНИЙ ===

def analyze_sheet_charts(sheet) -> List[Dict[str, Any]]:
    """
    Анализирует диаграммы на листе.
    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): Объект листа Excel.
    Returns:
        List[Dict[str, Any]]: Список словарей с информацией о диаграммах.
    """
    logger.debug(f"[ДИАГРАММЫ] Начало анализа диаграмм на листе '{sheet.title}'")
    charts_info = []
    try:
        if hasattr(sheet, '_charts'):
            charts = sheet._charts
            logger.debug(f"[ДИАГРАММЫ] Найдено {len(charts)} диаграмм на листе '{sheet.title}' (через _charts).")
            for i, chart in enumerate(charts):
                logger.debug(f"[ДИАГРАММЫ] Обработка диаграммы {i+1}/{len(charts)}")
                chart_data = extract_chart_data(chart, sheet) # Используем улучшенную функцию
                if chart_data:
                    charts_info.append(chart_data)
        else:
            logger.warning(f"[ДИАГРАММЫ] Атрибут _charts не найден у листа '{sheet.title}'.")
            
        logger.debug(f"[ДИАГРАММЫ] Анализ диаграмм на листе '{sheet.title}' завершен. Итого: {len(charts_info)} диаграмм.")
        return charts_info
    except Exception as e:
        logger.error(f"[ДИАГРАММЫ] Ошибка при анализе диаграмм на листе '{sheet.title}': {e}", exc_info=True)
        return []

# === НОВОЕ: Функция для анализа объединенных ячеек ===
def analyze_sheet_merged_cells(sheet) -> List[str]:
    """
    Анализирует объединенные ячейки на листе.
    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): Объект листа Excel.
    Returns:
        List[str]: Список строковых представлений диапазонов объединенных ячеек (например, 'A1:C3').
    """
    logger.debug(f"[ОБЪЕДИНЕННЫЕ ЯЧЕЙКИ] Начало анализа для листа '{sheet.title}'")
    merged_ranges = []
    try:
        merged_ranges = [str(mcr) for mcr in sheet.merged_cells.ranges]
        logger.debug(f"[ОБЪЕДИНЕННЫЕ ЯЧЕЙКИ] Найдено {len(merged_ranges)} объединенных диапазонов на листе '{sheet.title}'.")
    except Exception as e:
        logger.error(f"[ОБЪЕДИНЕННЫЕ ЯЧЕЙКИ] Ошибка при анализе для листа '{sheet.title}': {e}", exc_info=True)
    return merged_ranges
# === КОНЕЦ НОВОГО ===

def analyze_excel_file(file_path: str, sheet_names: Optional[List[str]] = None) -> Optional[Dict[str, Any]]:
    """
    Основная функция для анализа Excel файла.
    Args:
        file_path (str): Путь к Excel файлу.
        sheet_names (List[str], optional): Список имен листов для анализа. 
                                          Если None, анализируются все листы.
    Returns:
        Optional[Dict[str, Any]]: Словарь с результатами анализа или None в случае ошибки.
    """
    logger.info(f"Начало анализа Excel файла: {file_path}")
    try:
        if not Path(file_path).exists():
            logger.error(f"Файл не найден: {file_path}")
            return None

        documentation_template = load_documentation_template()
        
        documentation = documentation_template.copy()
        documentation["file_path"] = file_path
        documentation["analysis_timestamp"] = datetime.now().isoformat()
        
        documentation["sheets"] = {}

        # data_only=False для получения формул и стилей
        wb = load_workbook(file_path, data_only=False) 

        if sheet_names is None:
            sheet_names = wb.sheetnames
            
        logger.info(f"Будет проанализировано {len(sheet_names)} листов: {sheet_names}")

        for idx, sheet_name in enumerate(sheet_names):
            logger.info(f"Анализ листа {idx+1}/{len(sheet_names)}: '{sheet_name}'")
            if sheet_name not in wb.sheetnames:
                logger.warning(f"Лист '{sheet_name}' не найден в файле. Пропущен.")
                continue
                
            sheet = wb[sheet_name]
            
            # --- АНАЛИЗ ---
            sheet_structure = analyze_sheet_structure(sheet)
            sheet_raw_data = analyze_sheet_raw_data(sheet) 
            sheet_formulas = analyze_sheet_formulas(sheet)
            cross_sheet_refs = analyze_cross_sheet_references(sheet_formulas, sheet_name)
            sheet_charts = analyze_sheet_charts(sheet)
            sheet_styles = analyze_sheet_styles(sheet) # НОВОЕ
            sheet_merged_cells = analyze_sheet_merged_cells(sheet) # НОВОЕ
            
            # Используем словарь для хранения информации о листе
            sheet_info = {
                "name": sheet_name,
                "index": idx,
                "structure": sheet_structure,
                "raw_data": sheet_raw_data,
                "formulas": sheet_formulas,
                "cross_sheet_references": cross_sheet_refs,
                "charts": sheet_charts,
                "styled_ranges": sheet_styles, # НОВОЕ
                "merged_cells": sheet_merged_cells # НОВОЕ
            }
            
            documentation["sheets"][sheet_name] = sheet_info

        logger.info(f"Анализ Excel файла завершен: {file_path}")
        return documentation

    except Exception as e:
        logger.error(f"Ошибка при анализе Excel файла {file_path}: {e}", exc_info=True)
        return None

# - ТОЧКА ВХОДА ДЛЯ ТЕСТИРОВАНИЯ -
if __name__ == "__main__":
    print("--- ТЕСТ АНАЛИЗАТОРА ---")
    test_file = project_root / "data" / "samples" / "test_sample.xlsx"
    print(f"Путь к тестовому файлу: {test_file}")
    
    if test_file.exists():
        print(f"Тестовый файл найден: {test_file}")
        print("Начало анализа...")
        result = analyze_excel_file(str(test_file))
        if result:
            print("Анализ завершен успешно.")
            print(f"Файл: {result['file_path']}")
            print(f"Время анализа: {result['analysis_timestamp']}")
            print(f"Всего листов: {len(result['sheets'])}")
            if result['sheets']:
                first_sheet_name = list(result['sheets'].keys())[0]
                print(f"Первый лист: {first_sheet_name}")
                print(f" Структура: {result['sheets'][first_sheet_name]['structure'][:2]}...")
                print(f" Сырые данные (первые 2 строки): {result['sheets'][first_sheet_name]['raw_data']['rows'][:2]}...")
                print(f" Формулы: {result['sheets'][first_sheet_name]['formulas'][:1]}...")
                print(f" Диаграммы: {len(result['sheets'][first_sheet_name]['charts'])} шт.")
                print(f" Стили: {len(result['sheets'][first_sheet_name]['styled_ranges'])} записей.")
                print(f" Объединенные ячейки: {len(result['sheets'][first_sheet_name]['merged_cells'])} шт.")
        else:
            print("Анализ завершился с ошибкой.")
    else:
        print(f"Тестовый файл не найден: {test_file}")
