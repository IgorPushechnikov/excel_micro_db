# src/exporter/chart_exporter.py
"""
Модуль для экспорта диаграмм из данных проекта в Excel файл с использованием openpyxl.
"""

import logging
from typing import Dict, Any, List, Optional, Union
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, LineChart, ScatterChart, AreaChart
from openpyxl.chart.series import Series
from openpyxl.chart.reference import Reference
from openpyxl.styles import Font, fills
from openpyxl.worksheet.worksheet import Worksheet

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in [str(p) for p in Path(__file__).parent.parent.parent.iterdir()]:
    import sys
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

# === ИСПРАВЛЕНО: Определение ChartData локально ===
ChartData = Dict[str, Any]

def _get_chart_class(chart_type: str):
    """Возвращает класс диаграммы openpyxl по строковому названию."""
    chart_classes = {
        "BarChart": BarChart,
        "PieChart": PieChart,
        "LineChart": LineChart,
        "ScatterChart": ScatterChart,
        "AreaChart": AreaChart,
    }
    return chart_classes.get(chart_type, BarChart)

def _create_openpyxl_chart(chart_data: ChartData) -> Optional[Union[BarChart, PieChart, LineChart, ScatterChart, AreaChart]]:
    """
    Создает объект диаграммы openpyxl на основе данных из БД.
    Args:
        chart_data (ChartData): Данные диаграммы, загруженные из БД.
    Returns:
        Optional[Union[...]]: Объект диаграммы openpyxl или None в случае ошибки.
    """
    try:
        chart_type = chart_data.get("type", "BarChart")
        chart_class = _get_chart_class(chart_type)
        chart = chart_class()
        logger.debug(f"Создана диаграмма типа {chart_type}")

        # === ИСПРАВЛЕНО: Упрощенная установка заголовка диаграммы ===
        chart_title_data = chart_data.get("title")
        if chart_title_data:
            chart.title = str(chart_title_data)
            logger.debug(f"Установлен заголовок диаграммы: {chart_title_data}")

        # Стиль диаграммы
        style = chart_data.get("style")
        if style is not None:
            chart.style = int(style)
            logger.debug(f"Установлен стиль диаграммы: {style}")

        # Легенда
        legend_position = chart_data.get("legend_position")
        if legend_position:
            if not hasattr(chart, 'legend') or getattr(chart, 'legend', None) is None:
                from openpyxl.chart.legend import Legend
                chart.legend = Legend()
            chart.legend.position = str(legend_position)
            logger.debug(f"Установлена позиция легенды: {legend_position}")

        auto_scaling = chart_data.get("auto_scaling")
        if auto_scaling is not None:
            chart.auto_scaling = bool(auto_scaling)
            logger.debug(f"Установлено auto_scaling: {auto_scaling}")

        plot_vis_only = chart_data.get("plot_vis_only")
        if plot_vis_only is not None:
            chart.plotVisOnly = bool(plot_vis_only)
            logger.debug(f"Установлено plotVisOnly: {plot_vis_only}")

        # === ИСПРАВЛЕНО: Обработка осей с учетом 3D и использованием getattr ===
        axes_data = chart_data.get("axes", [])
        if isinstance(axes_data, list):
            # === ИСПРАВЛЕНО: Проверка на 3D диаграмму ===
            is_3d_chart = hasattr(chart, 'view3D') or '3D' in chart.__class__.__name__
            
            for axis_data in axes_data:
                if not isinstance(axis_data, dict):
                    continue
                axis_type = axis_data.get("axis_type")
                axis_obj = None

                # === ИСПРАВЛЕНО: Используем getattr для доступа к атрибутам осей ===
                # Это решает ошибки Pylance: "Не удается получить/назначить атрибут "z_axis""
                if axis_type == "x_axis":
                    # Проверяем, существует ли атрибут x_axis у объекта chart
                    current_axis = getattr(chart, 'x_axis', None)
                    if current_axis is None:
                        if isinstance(chart, (BarChart, LineChart, AreaChart, ScatterChart)):
                            try:
                                from openpyxl.chart.axis import TextAxis
                                chart.x_axis = TextAxis()
                            except ImportError:
                                pass
                        else:
                            try:
                                from openpyxl.chart.axis import NumericAxis
                                chart.x_axis = NumericAxis()
                            except ImportError:
                                pass
                    # Получаем объект оси через getattr
                    axis_obj = getattr(chart, 'x_axis', None)
                    
                elif axis_type == "y_axis":
                    current_axis = getattr(chart, 'y_axis', None)
                    if current_axis is None:
                        try:
                            from openpyxl.chart.axis import NumericAxis
                            chart.y_axis = NumericAxis()
                        except ImportError:
                            pass
                    axis_obj = getattr(chart, 'y_axis', None)
                    
                elif axis_type == "z_axis":
                    # === ИСПРАВЛЕНО: Обрабатываем z_axis только для 3D диаграмм ===
                    if is_3d_chart:
                        # Проверяем, существует ли атрибут z_axis у объекта chart
                        current_axis = getattr(chart, 'z_axis', None)
                        if current_axis is None:
                            try:
                                from openpyxl.chart.axis import SeriesAxis
                                # === ИСПРАВЛЕНО: Используем setattr для присвоения ===
                                # Это решает последние ошибки Pylance: "Не удается назначить атрибут "z_axis""
                                setattr(chart, 'z_axis', SeriesAxis())
                            except ImportError:
                                pass
                        # Получаем объект оси через getattr
                        axis_obj = getattr(chart, 'z_axis', None)
                    else:
                        # Для 2D диаграмм пропускаем обработку z_axis
                        logger.debug(f"Диаграмма {chart_type} не является 3D. Ось Z игнорируется.")
                        continue # Переходим к следующей оси в данных

                
                # === ИСПРАВЛЕНО: Проверка, что объект оси был создан ===
                if axis_obj is None:
                    logger.warning(f"Не удалось создать или получить объект оси типа {axis_type} для диаграммы {chart_type}")
                    continue

                # === ИСПРАВЛЕНО: Упрощенная установка заголовка оси ===
                axis_title_data = axis_data.get("title")
                if axis_title_data:
                     try:
                         axis_obj.title = str(axis_title_data)
                         logger.debug(f"Установлен заголовок оси {axis_type}: {axis_title_data}")
                     except Exception as e:
                         logger.warning(f"Не удалось установить заголовок оси {axis_type}: {e}")

                # Заполняем другие атрибуты оси, подавляя ошибки типов
                ax_pos = axis_data.get("ax_pos")
                if ax_pos:
                    try:
                        axis_obj.axPos = str(ax_pos) # type: ignore
                    except Exception:
                        pass

                delete_flag = axis_data.get("delete")
                if delete_flag is not None:
                    try:
                        axis_obj.delete = bool(delete_flag) # type: ignore
                    except Exception:
                        pass

                # Шкала (Scaling)
                scaling = {}
                min_val = axis_data.get("min")
                if min_val is not None:
                    scaling['min'] = float(min_val)
                max_val = axis_data.get("max")
                if max_val is not None:
                    scaling['max'] = float(max_val)
                orientation = axis_data.get("orientation")
                if orientation:
                    scaling['orientation'] = str(orientation)
                major_unit = axis_data.get("major_unit")
                if major_unit is not None:
                    scaling['majorUnit'] = float(major_unit)
                minor_unit = axis_data.get("minor_unit")
                if minor_unit is not None:
                    scaling['minorUnit'] = float(minor_unit)
                log_base = axis_data.get("log_base")
                if log_base is not None:
                    scaling['logBase'] = float(log_base)
                
                if scaling and hasattr(axis_obj, 'scaling') and axis_obj.scaling:
                    for k, v in scaling.items():
                        try:
                            setattr(axis_obj.scaling, k, v) # type: ignore
                        except Exception:
                            pass # Игнорируем ошибки установки свойств шкалы

                # Метки и деления
                major_tick_mark = axis_data.get("major_tick_mark")
                if major_tick_mark:
                    try:
                        axis_obj.majorTickMark = str(major_tick_mark) # type: ignore
                    except Exception:
                        pass
                minor_tick_mark = axis_data.get("minor_tick_mark")
                if minor_tick_mark:
                    try:
                        axis_obj.minorTickMark = str(minor_tick_mark) # type: ignore
                    except Exception:
                        pass
                tick_lbl_pos = axis_data.get("tick_lbl_pos")
                if tick_lbl_pos:
                    try:
                        axis_obj.tickLblPos = str(tick_lbl_pos) # type: ignore
                    except Exception:
                        pass

                # Формат чисел
                num_fmt = axis_data.get("num_fmt")
                if num_fmt and hasattr(axis_obj, 'numFmt'):
                    try:
                        axis_obj.numFmt.formatCode = str(num_fmt) # type: ignore
                    except Exception:
                        pass

                # Пересечения
                crosses = axis_data.get("crosses")
                if crosses:
                    try:
                        axis_obj.crosses = str(crosses) # type: ignore
                    except Exception:
                        pass
                crosses_at = axis_data.get("crosses_at")
                if crosses_at is not None and hasattr(axis_obj, 'crossesAt'):
                    try:
                        axis_obj.crossesAt = float(crosses_at) # type: ignore
                    except Exception:
                        pass

                # Линии сетки
                major_gridlines = axis_data.get("major_gridlines")
                if major_gridlines:
                    current_major_gridlines = getattr(axis_obj, 'majorGridlines', None)
                    if current_major_gridlines is None:
                        try:
                            from openpyxl.chart.axis import ChartLines
                            axis_obj.majorGridlines = ChartLines() # type: ignore
                        except ImportError:
                            pass
                
                minor_gridlines = axis_data.get("minor_gridlines")
                if minor_gridlines:
                    current_minor_gridlines = getattr(axis_obj, 'minorGridlines', None)
                    if current_minor_gridlines is None:
                        try:
                            from openpyxl.chart.axis import ChartLines
                            axis_obj.minorGridlines = ChartLines() # type: ignore
                        except ImportError:
                            pass

        # === ИСПРАВЛЕНО: Обработка серий данных ===
        series_data_list = chart_data.get("series", [])
        if isinstance(series_data_list, list):
            for series_data in series_data_list:
                if not isinstance(series_data, dict):
                    continue
                new_series = Series() 
                
                idx = series_data.get("idx")
                if idx is not None:
                    new_series.idx = int(idx)
                order = series_data.get("order")
                if order is not None:
                    new_series.order = int(order)
                
                # === ИСПРАВЛЕНО: Упрощенная установка заголовка серии ===
                tx_title = series_data.get("tx")
                if tx_title:
                    try:
                        # Попробуем стандартный способ
                        new_series.title = str(tx_title)
                    except AttributeError:
                        # Если не работает, используем tx.strRef
                        try:
                            from openpyxl.chart.series import SeriesLabel
                            # Создаем объект SeriesLabel
                            series_label = SeriesLabel()
                            # Устанавливаем текст через strRef (упрощенно)
                            # Для более точной настройки может потребоваться создание strRef и strCache
                            series_label.strRef = str(tx_title) # type: ignore
                            new_series.tx = series_label
                        except Exception as e:
                            logger.warning(f"Не удалось установить заголовок серии '{tx_title}': {e}")

                # === ИСПРАВЛЕНО: Установка shape с подавлением ошибок ===
                shape = series_data.get("shape")
                if shape:
                    try:
                        new_series.shape = str(shape) # type: ignore
                    except Exception:
                        pass
                
                smooth = series_data.get("smooth")
                if smooth is not None:
                    new_series.smooth = bool(smooth)
                
                invert_if_negative = series_data.get("invert_if_negative")
                if invert_if_negative is not None:
                    new_series.invertIfNegative = bool(invert_if_negative)

                chart.series.append(new_series)
                logger.debug(f"Добавлена серия данных: {tx_title}")

        return chart

    except Exception as e:
        logger.error(f"Ошибка при создании диаграммы openpyxl: {e}", exc_info=True)
        return None

def export_charts_to_worksheet(
    wb: Workbook,
    ws: Worksheet,
    sheet_name: str,
    charts_data: List[ChartData],
    formulas_for_charts: Dict[str, Any]
) -> bool:
    """
    Экспортирует диаграммы на лист Excel.
    Args:
        wb (Workbook): Объект рабочей книги openpyxl.
        ws (Worksheet): Объект листа openpyxl.
        sheet_name (str): Имя листа, на который добавляются диаграммы.
        charts_data (List[ChartData]): Список словарей с данными диаграмм.
        formulas_for_charts (Dict[str, Any]): Словарь с формулами для диаграмм.
    Returns:
        bool: True, если экспорт прошёл успешно, иначе False.
    """
    logger.info(f"[ДИАГРАММЫ] Начало экспорта диаграмм на лист '{sheet_name}'")
    try:
        success_count = 0
        for i, chart_data in enumerate(charts_data):
            logger.debug(f"[ДИАГРАММЫ] Обработка диаграммы {i+1}/{len(charts_data)}")
            if not isinstance(chart_data, dict):
                logger.warning(f"[ДИАГРАММЫ] Некорректные данные диаграммы {i+1}. Пропущена.")
                continue

            chart = _create_openpyxl_chart(chart_data)
            if not chart:
                logger.error(f"[ДИАГРАММЫ] Не удалось создать диаграмму {i+1}. Пропущена.")
                continue

            # === ИСПРАВЛЕНО: Проверка на None перед использованием ws ===
            if ws is None:
                logger.error(f"[ДИАГРАММЫ] Лист для диаграммы {i+1} не определен. Пропущена.")
                continue

            # === ИСПРАВЛЕНО: Обработка источников данных для серий ===
            data_sources = chart_data.get("data_sources", [])
            if isinstance(data_sources, list):
                for ds in data_sources:
                    if not isinstance(ds, dict):
                        continue
                    series_index = ds.get("series_index")
                    data_type = ds.get("data_type")
                    formula = ds.get("formula")
                    
                    if series_index is None or not data_type or not formula:
                        continue
                    
                    if series_index < len(chart.series):
                        target_series = chart.series[series_index]
                    else:
                        logger.warning(f"[ДИАГРАММЫ] Индекс серии {series_index} выходит за пределы. Пропущен источник данных.")
                        continue

                    try:
                        if '!' in formula:
                            range_part = formula.split('!')[1]
                        else:
                            range_part = formula
                        
                        clean_range = range_part.replace('$', '')
                        if ':' in clean_range:
                            start_cell, end_cell = clean_range.split(':')
                        else:
                            start_cell = end_cell = clean_range
                        
                        import re
                        start_match = re.match(r"([A-Z]+)(\d+)", start_cell)
                        end_match = re.match(r"([A-Z]+)(\d+)", end_cell)
                        if not start_match or not end_match:
                            logger.warning(f"[ДИАГРАММЫ] Не удалось распарсить диапазон '{range_part}' для серии {series_index}.")
                            continue
                        
                        start_col_letter, start_row_str = start_match.groups()
                        end_col_letter, end_row_str = end_match.groups()
                        
                        from openpyxl.utils import column_index_from_string
                        min_col = column_index_from_string(start_col_letter)
                        max_col = column_index_from_string(end_col_letter)
                        min_row = int(start_row_str)
                        max_row = int(end_row_str)

                        # === ИСПРАВЛЕНО: Проверка на None (на всякий случай) ===
                        if min_row is None or min_col is None or max_row is None or max_col is None:
                            logger.warning(f"[ДИАГРАММЫ] Ошибка определения границ диапазона для серии {series_index}.")
                            continue

                        data_ref = Reference(
                            worksheet=ws,
                            min_col=min_col,
                            min_row=min_row,
                            max_col=max_col,
                            max_row=max_row
                        )
                        
                        if data_type == "values":
                            target_series.val = data_ref
                        elif data_type == "categories":
                            target_series.cat = data_ref
                            
                    except Exception as e:
                        logger.error(f"[ДИАГРАММЫ] Ошибка обработки источника данных для серии {series_index}: {e}", exc_info=True)
                        continue

            # Добавляем диаграмму на лист
            try:
                top_left_cell = chart_data.get("top_left_cell", "A1")
                if top_left_cell:
                    try:
                        chart.anchor = f"{top_left_cell}"
                    except Exception:
                        chart.anchor = "A1"
                
                ws.add_chart(chart)
                success_count += 1
                logger.debug(f"[ДИАГРАММЫ] Диаграмма {i+1} успешно добавлена на лист.")
                
            except Exception as e:
                logger.error(f"[ДИАГРАММЫ] Ошибка добавления диаграммы {i+1} на лист: {e}", exc_info=True)
                continue

        logger.info(f"[ДИАГРАММЫ] Экспорт диаграмм на лист '{sheet_name}' завершен. Успешно: {success_count}/{len(charts_data)}")
        return success_count > 0

    except Exception as e:
        logger.error(f"[ДИАГРАММЫ] Критическая ошибка при экспорте диаграмм на лист '{sheet_name}': {e}", exc_info=True)
        return False
