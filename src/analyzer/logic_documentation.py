# src/analyzer/logic_documentation.py

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
# Импортируем Cell напрямую
from openpyxl.cell.cell import Cell
# Импортируем модуль merge, не конкретный класс
import openpyxl.worksheet.merge
# Для аннотаций типов
from typing import Dict, Any, List, Optional, Union
import logging
import json
# Импортируем logger из utils
from src.utils.logger import get_logger

logger = get_logger(__name__)

# --- Вспомогательные функции для сериализации сложных объектов ---

# Используем Union для аннотации типа параметра cell
# Используем строку для аннотации типа, чтобы избежать проблем импорта
# И предполагаем, что MergedCell, если существует, совместим с Cell для наших целей
# Если MergedCell не существует или не импортируется, мы просто аннотируем как Cell
# Это снимет ошибку Pylance, так как Cell - точно существующий тип
def _serialize_style(cell: Cell) -> Dict[str, Any]:
# Или, если вы хотите быть максимально точным (но рискуете с ошибкой Pylance)
# def _serialize_style(cell: Union[Cell, 'openpyxl.worksheet.merge.MergedCell']) -> Dict[str, Any]:
# Но если MergedCell не находится, лучше упростить до Cell
# def _serialize_style(cell: Cell) -> Dict[str, Any]:

    """
    Сериализует атрибуты стиля ячейки openpyxl в словарь.
    Этот словарь будет сериализован в JSON в storage/styles.py.
    Структура должна соответствовать ожиданиям _convert_style_to_xlsxwriter_format
    в excel_exporter.py.
    """
    # ... (остальная реализация функции остается прежней)
    # Поскольку MergedCell наследуется от Cell (если существует) или ведет себя как Cell,
    # доступ к атрибутам типа cell.font, cell.fill должен работать.
    style_dict = {}

    # --- Шрифт ---
    if hasattr(cell, 'font') and cell.font:
        font_dict = {}
        if cell.font.name: font_dict['name'] = cell.font.name
        if cell.font.sz: font_dict['sz'] = cell.font.sz # Размер
        if cell.font.b: font_dict['b'] = cell.font.b # Жирный
        if cell.font.i: font_dict['i'] = cell.font.i # Курсив
        if cell.font.color and cell.font.color.rgb:
            font_dict['color'] = {'rgb': cell.font.color.rgb}
        # ... другие атрибуты шрифта (underline, vertAlign и т.д.)
        if font_dict:
            style_dict['font'] = font_dict

    # --- Заливка ---
    if hasattr(cell, 'fill') and cell.fill:
        fill_dict = {}
        # openpyxl.fill.PatternFill или openpyxl.fill.GradientFill
        if hasattr(cell.fill, 'patternType'):
             fill_dict['patternType'] = cell.fill.patternType
        if hasattr(cell.fill, 'fgColor') and cell.fill.fgColor and cell.fill.fgColor.rgb:
             fill_dict['fgColor'] = {'rgb': cell.fill.fgColor.rgb}
        if hasattr(cell.fill, 'bgColor') and cell.fill.bgColor and cell.fill.bgColor.rgb:
             fill_dict['bgColor'] = {'rgb': cell.fill.bgColor.rgb}
        # ... обработка GradientFill
        if fill_dict:
             style_dict['fill'] = fill_dict

    # --- Границы ---
    if hasattr(cell, 'border') and cell.border:
        border_dict = {}
        # openpyxl.styles.borders.Border
        for side_name in ['left', 'right', 'top', 'bottom']: # 'diagonal' ?
            side_obj = getattr(cell.border, side_name, None)
            if side_obj and (side_obj.style or (side_obj.color and side_obj.color.rgb)):
                side_dict = {}
                if side_obj.style: side_dict['style'] = side_obj.style
                if side_obj.color and side_obj.color.rgb:
                    side_dict['color'] = {'rgb': side_obj.color.rgb}
                border_dict[side_name] = side_dict
        if border_dict:
            style_dict['border'] = border_dict

    # --- Выравнивание ---
    if hasattr(cell, 'alignment') and cell.alignment:
        alignment_dict = {}
        # openpyxl.styles.alignment.Alignment
        if cell.alignment.horizontal: alignment_dict['horizontal'] = cell.alignment.horizontal
        if cell.alignment.vertical: alignment_dict['vertical'] = cell.alignment.vertical
        if cell.alignment.wrapText is not None: alignment_dict['wrapText'] = cell.alignment.wrapText
        if cell.alignment.textRotation is not None: alignment_dict['textRotation'] = cell.alignment.textRotation
        # ... другие атрибуты (shrinkToFit, indent и т.д.)
        if alignment_dict:
            style_dict['alignment'] = alignment_dict

    # --- Числовой формат ---
    if hasattr(cell, 'number_format') and cell.number_format:
        style_dict['number_format'] = cell.number_format

    # --- Защита ---
    if hasattr(cell, 'protection') and cell.protection:
        protection_dict = {}
        # openpyxl.styles.protection.Protection
        if cell.protection.locked is not None: protection_dict['locked'] = cell.protection.locked
        if cell.protection.hidden is not None: protection_dict['hidden'] = cell.protection.hidden
        if protection_dict:
            style_dict['protection'] = protection_dict
            
    # logger.debug(f"Сериализован стиль для ячейки {cell.coordinate}: {list(style_dict.keys())}")
    return style_dict

def _serialize_chart(chart_obj) -> Dict[str, Any]:
    """
    Сериализует объект диаграммы openpyxl.
    В данном примере мы будем сохранять XML-представление диаграммы,
    которое можно будет использовать при экспорте.
    """
    # openpyxl.chart.chart.ChartBase и его подклассы имеют атрибут _chart_space
    # который содержит XML-элемент.
    # Также можно попробовать chart_obj.to_tree() или chart_obj._write()
    # Но самый надежный способ - получить XML напрямую.
    
    try:
        # Альтернатива: сохранить как словарь с ключевыми атрибутами
        # Это может быть проще для десериализации, но сложнее для точного воссоздания
        chart_data = {}
        chart_data['type'] = type(chart_obj).__name__ # 'BarChart', 'LineChart' и т.д.
        
        # Пример: сохранение ссылок на данные
        if hasattr(chart_obj, 'ser') and chart_obj.ser:
            series_data = []
            for idx, s in enumerate(chart_obj.ser):
                ser_dict = {}
                # Сохраняем адреса диапазонов данных
                if hasattr(s, 'val') and s.val and hasattr(s.val, 'numRef') and s.val.numRef:
                    ser_dict['val_range'] = s.val.numRef.f # Строка формулы диапазона значений
                if hasattr(s, 'cat') and s.cat and hasattr(s.cat, 'strRef') and s.cat.strRef:
                    ser_dict['cat_range'] = s.cat.strRef.f # Строка формулы диапазона категорий
                elif hasattr(s, 'cat') and s.cat and hasattr(s.cat, 'numRef') and s.cat.numRef:
                    ser_dict['cat_range'] = s.cat.numRef.f
                # ... другие атрибуты серии (название, цвета и т.д.)
                series_data.append(ser_dict)
            chart_data['series'] = series_data
            
        # Пример: сохранение заголовка
        if hasattr(chart_obj, 'title') and chart_obj.title:
             if hasattr(chart_obj.title, 'tx') and chart_obj.title.tx:
                 # Попробуем обойтись без прямого импорта RichText
                 # Проверим, есть ли атрибут rich и он имеет атрибуты p, r, t
                 if hasattr(chart_obj.title.tx, 'rich') and chart_obj.title.tx.rich:
                     # Получить текст из rich text
                     # from openpyxl.drawing.text import RichText - уже импортирован
                     # Проверим тип объекта
                     # if isinstance(chart_obj.title.tx.rich, RichText): # Убираем прямую проверку типа
                     try:
                         # Попробуем получить текст "в лоб"
                         # chart_data['title'] = chart_obj.title.tx.rich ... (нужно извлечь текст)
                         # Упрощение: берем первый run
                         if hasattr(chart_obj.title.tx.rich, 'p') and chart_obj.title.tx.rich.p and len(chart_obj.title.tx.rich.p) > 0:
                             first_p = chart_obj.title.tx.rich.p[0]
                             if first_p and hasattr(first_p, 'r') and first_p.r and len(first_p.r) > 0:
                                 # Проверим, есть ли атрибут t у run
                                 if hasattr(first_p.r[0], 't'):
                                     chart_data['title'] = first_p.r[0].t
                     except AttributeError as ae:
                         logger.debug(f"Не удалось извлечь заголовок из rich text: {ae}")
                 elif hasattr(chart_obj.title.tx, 'strRef') and chart_obj.title.tx.strRef:
                     chart_data['title_ref'] = chart_obj.title.tx.strRef.f # Ссылка на ячейку с заголовком
                     
        # ... другие атрибуты диаграммы (legend, axId, plotArea и т.д.)
        
        logger.debug(f"Сериализована диаграмма типа {chart_data.get('type', 'Unknown')}")
        return chart_data
        
    except Exception as e:
        logger.error(f"Ошибка при сериализации диаграммы: {e}", exc_info=True)
        # Возвращаем пустой словарь или None в случае ошибки
        return {}

# ... (остальной код analyze_excel_file остается в основном без изменений, 
#     за исключением аннотации типов и обработки ошибок)

# --- Основная функция анализа ---
# ... (внутри analyze_excel_file, в цикле по строкам и ячейкам)
# for row in sheet.iter_rows(values_only=False):
#     for cell in row:
#         # Проверяем, есть ли у ячейки не-дефолтный стиль
#         # Это можно сделать, сравнивая с дефолтным стилем, но проще сериализовать всегда
#         # и потом в storage/styles.py решать, нужно ли его сохранять.
#         
#         # Передаем cell в _serialize_style, который теперь принимает Union[Cell, MergedCell]
#         style_dict = _serialize_style(cell)
#         if style_dict: # Если стиль не пустой
#             style_json = json.dumps(style_dict, sort_keys=True) # Сериализуем в JSON и сортируем ключи для надежного хеширования
#             coord = cell.coordinate
#             
#             if style_json in style_ranges_map:
#                 style_ranges_map[style_json].append(coord)
#             else:
#                 style_ranges_map[style_json] = [coord]

# --- 4. Извлечение диаграмм ---
# ... (внутри analyze_excel_file)
# logger.debug(f"Извлечение диаграмм с листа '{sheet_name}'...")
# # Диаграммы находятся в sheet._charts
# # Оборачиваем в try...except, так как доступ к _charts может быть ненадежным
# # Используем # type: ignore для подавления предупреждения Pylance
# try:
#     for chart_obj in sheet._charts: # type: ignore[attr-defined]
#          chart_data = _serialize_chart(chart_obj)
#          if chart_data:
#              # storage ожидает 'chart_data' как сериализованный объект
#              sheet_data["charts"].append({
#                  "chart_data": chart_data # Это будет словарь, storage должен его сериализовать при сохранении
#                  # Если нужно сохранить как JSON сразу:
#                  # "chart_data": json.dumps(chart_data, ensure_ascii=False)
#              })
# except AttributeError as ae:
#     logger.warning(f"Не удалось получить доступ к диаграммам листа '{sheet_name}': {ae}")
# except Exception as e:
#     logger.error(f"Ошибка при извлечении диаграмм с листа '{sheet_name}': {e}", exc_info=True)
