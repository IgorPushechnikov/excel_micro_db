# src/exporter/chart_exporter.py
"""Модуль для экспорта диаграмм листа Excel."""
import sys
from pathlib import Path
from typing import Dict, Any, List, Optional, Type, Union # Добавлен Union
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
# --- ИСПРАВЛЕННЫЕ ИМПОРТЫ ---
# Импортируем конкретные типы диаграмм из openpyxl.chart (__init__.py)
from openpyxl.chart import (
    BarChart, LineChart, PieChart, AreaChart, ScatterChart, PieChart3D
)
# Импортируем ChartBase напрямую из модуля, где он определен
from openpyxl.chart._chart import ChartBase
# Импортируем Series и Reference
from openpyxl.chart.series import Series
from openpyxl.chart.reference import Reference
# --- КОНЕЦ ИСПРАВЛЕННЫХ ИМПОРТОВ ---
# Импортируем Axis, если нужно работать с осями
# from openpyxl.chart.axis import Axis # Раскомментируйте, если будете использовать

# Добавляем корень проекта в путь поиска модулей
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

# Карта типов диаграмм (теперь использует корректно импортированные классы)
CHART_TYPE_MAP: Dict[str, Type[ChartBase]] = {
    "BarChart": BarChart,
    "LineChart": LineChart,
    "PieChart": PieChart,
    "AreaChart": AreaChart,
    "ScatterChart": ScatterChart,
    "PieChart3D": PieChart3D,
    # Добавить другие поддерживаемые типы при необходимости
}


def _get_chart_class_by_type_name(chart_type_name: str) -> Optional[Type[ChartBase]]:
    """Получает класс диаграммы openpyxl по её строковому имени."""
    chart_class = CHART_TYPE_MAP.get(chart_type_name)
    if not chart_class:
        logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Тип диаграммы '{chart_type_name}' не поддерживается.")
    return chart_class

def _parse_range_string(range_str: str):
    """
    Парсит строку диапазона, например, "Sheet1!$A$1:$B$10" или "$A$1:$B$10".
    Возвращает кортеж: (sheet_name_or_none, min_col, min_row, max_col, max_row)
    """
    try:
        if '!' in range_str:
            sheet_part, cells_part = range_str.split('!', 1)
        else:
            sheet_part = None
            cells_part = range_str

        clean_cells = cells_part.replace('$', '')
        if ':' in clean_cells:
            start_cell, end_cell = clean_cells.split(':')
        else:
            start_cell = end_cell = clean_cells

        start_col_letter = ''.join(filter(str.isalpha, start_cell))
        start_row = int(''.join(filter(str.isdigit, start_cell)))
        end_col_letter = ''.join(filter(str.isalpha, end_cell))
        end_row = int(''.join(filter(str.isdigit, end_cell)))

        # openpyxl.utils использует 1-индексацию, как Excel
        min_col = ord(start_col_letter.upper()) - ord('A') + 1
        max_col = ord(end_col_letter.upper()) - ord('A') + 1
        min_row = start_row
        max_row = end_row

        return sheet_part, min_col, min_row, max_col, max_row
    except Exception as e:
        logger.error(f"[ЭКСПОРТ_ДИАГРАММ] Ошибка парсинга диапазона '{range_str}': {e}")
        # Возвращаем None для всех значений в случае ошибки
        return None, None, None, None, None

def _apply_chart_attributes(chart_obj: ChartBase, chart_info: Dict[str, Any]) -> None:
    """Применяет атрибуты диаграммы из словаря."""
    logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Применение атрибутов диаграммы: {chart_info.get('type')}")
    try:
        # Установка заголовка
        if chart_info.get("title") is not None:
            chart_obj.title = chart_info["title"]
        
        # Установка стиля
        if chart_info.get("style") is not None:
            chart_obj.style = chart_info["style"]
            
        # Установка положения легенды
        if chart_info.get("legend_position") is not None:
            if chart_obj.legend:
                 chart_obj.legend.position = chart_info["legend_position"]
            else:
                 logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] У диаграммы нет объекта legend для установки позиции {chart_info['legend_position']}")
                 
        # TODO: Добавить другие атрибуты диаграммы при необходимости
        # auto_scaling, plot_vis_only и т.д.
        
    except Exception as e:
        logger.error(f"[ЭКСПОРТ_ДИАГРАММ] Ошибка применения атрибутов диаграммы: {e}", exc_info=True)

# def _apply_axis_attributes(axis_obj: 'Axis', axis_info: Dict[str, Any]) -> None:
#     """Применяет атрибуты оси из словаря."""
#     # TODO: Реализовать применение атрибутов оси
#     pass

def _create_series_from_data(
    series_info: Dict[str, Any], 
    data_sources_data: List[Dict[str, Any]], 
    workbook: Workbook
) -> Optional[Series]:
    """
    Создает объект Series на основе информации из БД.
    Args:
        series_info (Dict): Информация о серии из БД (chart_series).
        data_sources_data (List[Dict]): Все источники данных диаграммы (chart_data_sources).
        workbook (Workbook): Рабочая книга для поиска листов источников.
    Returns:
        Optional[Series]: Созданный объект Series или None в случае ошибки.
    """
    logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Создание серии: {series_info}")
    try:
        series_idx = series_info.get("idx")
        series_tx = series_info.get("tx", f"Series {series_idx}")
        
        # 1. Найти источник данных значений (values) для этой серии
        values_ds_info = None
        categories_ds_info = None
        for ds_info in data_sources_data:
            # Связываем по series_index из chart_data_sources с idx из chart_series
            if ds_info.get("series_index") == series_idx and ds_info.get("data_type") == "values":
                values_ds_info = ds_info
            elif ds_info.get("series_index") == series_idx and ds_info.get("data_type") == "categories":
                 categories_ds_info = ds_info
        
        if not values_ds_info:
            logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Нет источника данных 'values' для серии idx={series_idx}")
            return None
            
        values_formula = values_ds_info.get("formula")
        if not values_formula:
             logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Пустая формула для источника 'values' серии idx={series_idx}")
             return None

        # 2. Парсим формулу диапазона значений
        sheet_part, min_col, min_row, max_col, max_row = _parse_range_string(values_formula)
        if min_col is None:
            logger.error(f"[ЭКСПОРТ_ДИАГРАММ] Ошибка парсинга формулы диапазона значений для серии idx={series_idx}")
            return None
            
        # 3. Определяем лист источника данных
        source_ws_name = sheet_part.strip("'") if sheet_part else workbook.active.title
        try:
            source_ws = workbook[source_ws_name]
            logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Найден лист источника данных: '{source_ws_name}' для серии idx={series_idx}")
        except KeyError:
            logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Лист источника данных '{source_ws_name}' не найден. Используется активный лист '{workbook.active.title}'.")
            source_ws = workbook.active # fallback

        # 4. Создаем Reference для значений
        values_ref = Reference(source_ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
        logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Создан Reference для значений серии idx={series_idx}: {values_formula}")

        # 5. Создаем объект Series
        # openpyxl.chart.series.Series(values, xvalues=None, zvalues=None, title=None, title_from_data=False)
        # Передаем title как строку
        series_obj = Series(values=values_ref, title=series_tx) 
        logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Создана серия idx={series_idx} с заголовком '{series_tx}'")

        # 6. Устанавливаем категории (xvalues), если они есть
        if categories_ds_info:
            categories_formula = categories_ds_info.get("formula")
            if categories_formula:
                cat_sheet_part, cat_min_col, cat_min_row, cat_max_col, cat_max_row = _parse_range_string(categories_formula)
                if cat_min_col is not None:
                    cat_source_ws_name = cat_sheet_part.strip("'") if cat_sheet_part else workbook.active.title
                    try:
                        cat_source_ws = workbook[cat_source_ws_name]
                        logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Найден лист источника категорий: '{cat_source_ws_name}' для серии idx={series_idx}")
                    except KeyError:
                        logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Лист источника категорий '{cat_source_ws_name}' не найден. Используется активный лист '{workbook.active.title}'.")
                        cat_source_ws = workbook.active # fallback
                        
                    categories_ref = Reference(cat_source_ws, min_col=cat_min_col, min_row=cat_min_row, max_col=cat_max_col, max_row=cat_max_row)
                    series_obj.xvalues = categories_ref
                    logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Категории установлены для серии idx={series_idx}")
                else:
                     logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Ошибка парсинга формулы категорий для серии idx={series_idx}")
        else:
             logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Нет источника данных 'categories' для серии idx={series_idx}")
             
        # TODO: Применить другие атрибуты серии (shape, smooth, invert_if_negative) если нужно
        # series_obj.shape = series_info.get("shape", ...)
        # series_obj.smooth = bool(series_info.get("smooth", False))
        # series_obj.invertIfNegative = bool(series_info.get("invert_if_negative", False))

        return series_obj

    except Exception as e:
        logger.error(f"[ЭКСПОРТ_ДИАГРАММ] Ошибка создания Series для серии idx={series_info.get('idx', 'unknown')}: {e}", exc_info=True)
        return None

def export_sheet_charts(ws: Worksheet, charts_data: List[Dict[str, Any]], workbook: Workbook) -> None:
    """
    Экспортирует диаграммы листа.
    Args:
        ws (Worksheet): Лист Excel, на который добавляются диаграммы.
        charts_data (List[Dict[str, Any]]): Список словарей с информацией о диаграммах.
        workbook (Workbook): Рабочая книга.
    """
    logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Начало экспорта диаграмм для листа '{ws.title}'")
    logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Получено {len(charts_data)} диаграмм.")

    if not charts_data:
        logger.debug("[ЭКСПОРТ_ДИАГРАММ] Нет диаграмм для экспорта.")
        return

    for i, chart_info_dict in enumerate(charts_data):
        logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Обработка диаграммы {i+1}: {chart_info_dict.get('type', 'UnknownType')}")
        try:
            # 1. Получаем тип диаграммы
            chart_type_name = chart_info_dict.get("type")
            if not chart_type_name:
                logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Пропущена диаграмма {i+1} из-за отсутствия типа.")
                continue

            # 2. Получаем класс диаграммы
            chart_class = _get_chart_class_by_type_name(chart_type_name)
            if not chart_class:
                logger.warning(f"[ЭКСПОРТ_ДИАГРАММ] Тип диаграммы '{chart_type_name}' не поддерживается. Пропущена диаграмма {i+1}.")
                continue

            # 3. Создаем новый экземпляр диаграммы
            chart_object: ChartBase = chart_class() # Аннотация типа для ясности
            logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Создана новая диаграмма типа {chart_type_name}")

            # 4. Применяем атрибуты диаграммы
            _apply_chart_attributes(chart_object, chart_info_dict)

            # 5. Устанавливаем позицию диаграммы
            top_left_cell = chart_info_dict.get("top_left_cell", "A1")
            chart_object.anchor = top_left_cell
            logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Установлена позиция диаграммы: {top_left_cell}")

            # 6. Обрабатываем серии данных
            # Предполагаем, что данные о сериях и источниках приходят корректно из БД
            series_data_list = chart_info_dict.get("series", [])
            data_sources_data_list = chart_info_dict.get("data_sources", [])
            
            # Сортируем серии по их order для правильного добавления
            sorted_series_data = sorted(series_data_list, key=lambda x: x.get("order", x.get("idx", 0)))

            for series_info in sorted_series_data:
                series_obj = _create_series_from_data(series_info, data_sources_data_list, workbook)
                if series_obj:
                    try:
                        chart_object.series.append(series_obj)
                        logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Серия idx={series_info.get('idx')} добавлена в диаграмму")
                    except Exception as e:
                        logger.error(f"[ЭКСПОРТ_ДИАГРАММ] Ошибка добавления серии idx={series_info.get('idx')} в диаграмму: {e}", exc_info=True)

            # TODO: Применить атрибуты осей, если они сохраняются и передаются корректно
            # axes_data_list = chart_info_dict.get("axes", [])
            # for axis_info in axes_data_list:
            #     # Логика применения атрибутов оси
            #     pass

            # 7. Добавляем диаграмму на лист
            ws.add_chart(chart_object, top_left_cell)
            logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] Диаграмма {i+1} ({chart_type_name}) с данными добавлена на лист в позицию {top_left_cell}.")

        except Exception as e:
            logger.error(f"[ЭКСПОРТ_ДИАГРАММ] Ошибка обработки диаграммы {i+1}: {e}", exc_info=True)

    logger.debug(f"[ЭКСПОРТ_ДИАГРАММ] === Конец экспорта диаграмм для листа '{ws.title}' ===")
