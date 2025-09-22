# src/storage/charts.py
"""
Модуль для работы с диаграммами в хранилище проекта Excel Micro DB.
"""
import sqlite3
import logging
from typing import List, Dict, Any

# from src.storage.schema import ... # Если потребуются какие-либо константы

logger = logging.getLogger(__name__)

def save_sheet_charts(connection: sqlite3.Connection, sheet_id: int, charts_data: List[Dict[str, Any]]) -> bool:
    """
    Сохраняет диаграммы для листа в структурированном виде.
    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа.
        charts_data (List[Dict[str, Any]]): Список словарей с данными диаграмм.
    Returns:
        bool: True, если диаграммы сохранены успешно, False в случае ошибки.
    """
    if not connection:
        logger.error("Получено пустое соединение для сохранения диаграмм.")
        return False
    try:
        cursor = connection.cursor()

        # Сначала удаляем старые диаграммы и их источники для этого листа
        # Удаляем источники данных (из-за внешнего ключа)
        chart_ids_to_delete = []
        cursor.execute("SELECT id FROM charts WHERE sheet_id = ?", (sheet_id,))
        for row in cursor.fetchall():
            chart_ids_to_delete.append(row[0])
        if chart_ids_to_delete:
            placeholders = ','.join('?' * len(chart_ids_to_delete))
            # Удаляем сначала серии, потом источники, потом диаграммы
            cursor.execute(f"SELECT id FROM chart_series WHERE chart_id IN ({placeholders})", chart_ids_to_delete)
            series_ids_to_delete = [r[0] for r in cursor.fetchall()]
            if series_ids_to_delete:
                series_placeholders = ','.join('?' * len(series_ids_to_delete))
                cursor.execute(f"DELETE FROM chart_data_sources WHERE series_id IN ({series_placeholders})", series_ids_to_delete)
            cursor.execute(f"DELETE FROM chart_series WHERE chart_id IN ({placeholders})", chart_ids_to_delete)
            cursor.execute(f"DELETE FROM chart_axes WHERE chart_id IN ({placeholders})", chart_ids_to_delete)
            cursor.execute(f"DELETE FROM charts WHERE sheet_id = ?", (sheet_id,))
        logger.debug(f"Удалены старые диаграммы для листа ID {sheet_id}.")

        for chart_info in charts_data:
            # Извлекаем основные атрибуты диаграммы
            chart_type = chart_info.get("type", "")
            chart_title = chart_info.get("title", "")
            top_left_cell = chart_info.get("top_left_cell", "")
            width = chart_info.get("width")
            height = chart_info.get("height")
            chart_style = chart_info.get("style")
            legend_position = chart_info.get("legend_position")
            auto_scaling = chart_info.get("auto_scaling")
            plot_vis_only = chart_info.get("plot_vis_only")
            if not chart_type:
                logger.warning(f"Пропущена диаграмма без типа: {chart_info}")
                continue
            # Вставляем запись о диаграмме
            # === ИСПРАВЛЕНО: Убраны кавычки с "order" и "references" ===
            cursor.execute('''
                INSERT INTO charts (sheet_id, type, title, top_left_cell, width, height, style, legend_position, auto_scaling, plot_vis_only)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (sheet_id, chart_type, chart_title, top_left_cell, width, height, chart_style, legend_position, auto_scaling, plot_vis_only))
            # ========================================================
            chart_id = cursor.lastrowid
            logger.debug(f"Создана запись диаграммы ID {chart_id} типа {chart_type}")
            # Вставляем оси диаграммы
            axes_data = chart_info.get("axes", [])
            for axis_info in axes_data:
                axis_type = axis_info.get("axis_type")
                ax_id = axis_info.get("ax_id")
                ax_pos = axis_info.get("ax_pos")
                delete_axis = axis_info.get("delete") # <-- ИСПРАВЛЕНО: delete_axis -> delete
                axis_title = axis_info.get("title") # <-- ИСПРАВЛЕНО: axis_title -> title
                num_fmt = axis_info.get("num_fmt") # <-- ИСПРАВЛЕНО: number_format -> num_fmt
                major_tick_mark = axis_info.get("major_tick_mark")
                minor_tick_mark = axis_info.get("minor_tick_mark")
                tick_lbl_pos = axis_info.get("tick_lbl_pos")
                crosses = axis_info.get("crosses")
                crosses_at = axis_info.get("crosses_at")
                major_unit = axis_info.get("major_unit")
                minor_unit = axis_info.get("minor_unit")
                min_val = axis_info.get("min")
                max_val = axis_info.get("max")
                orientation = axis_info.get("orientation")
                log_base = axis_info.get("log_base")
                major_gridlines = axis_info.get("major_gridlines")
                minor_gridlines = axis_info.get("minor_gridlines")
                # === ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py ===
                cursor.execute('''
                    INSERT INTO chart_axes (
                        chart_id, axis_type, ax_id, ax_pos, delete_axis, title, num_fmt,
                        major_tick_mark, minor_tick_mark, tick_lbl_pos, crosses, crosses_at,
                        major_unit, minor_unit, min, max, orientation, log_base,
                        major_gridlines, minor_gridlines
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    chart_id, axis_type, ax_id, ax_pos, delete_axis, axis_title, num_fmt,
                    major_tick_mark, minor_tick_mark, tick_lbl_pos, crosses, crosses_at,
                    major_unit, minor_unit, min_val, max_val, orientation, log_base,
                    major_gridlines, minor_gridlines
                ))
                # =========================================================
                logger.debug(f"  Добавлена ось {axis_type} для диаграммы ID {chart_id}")
            # Вставляем серии данных для этой диаграммы
            series_data = chart_info.get("series", [])
            data_sources_data = chart_info.get("data_sources", [])
            for series_info in series_data:
                # === ИСПРАВЛЕНО: Убраны кавычки с "order" ===
                series_idx = series_info.get("idx")
                series_order = series_info.get("order") # <-- ИСПРАВЛЕНО: "order" -> order
                series_tx = series_info.get("tx")
                shape = series_info.get("shape")
                smooth = series_info.get("smooth")
                invert_if_negative = series_info.get("invert_if_negative")
                # === ИСПРАВЛЕНО: Убраны кавычки с "order" ===
                cursor.execute('''
                    INSERT INTO chart_series (chart_id, idx, "order", tx, shape, smooth, invert_if_negative)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (chart_id, series_idx, series_order, series_tx, shape, smooth, invert_if_negative))
                # =================================================
                series_id = cursor.lastrowid
                logger.debug(f"  Добавлена серия ID {series_id} (idx {series_idx}) для диаграммы ID {chart_id}")
                # Вставляем источники данных, связанные с этой серией
                for ds_info in data_sources_data:
                    if ds_info.get("series_index") == series_idx: # Связываем по индексу серии
                        # === ИСПРАВЛЕНО: source_type -> data_type ===
                        data_type = ds_info.get("data_type") # <-- ИСПРАВЛЕНО
                        formula = ds_info.get("formula")
                        # === ИСПРАВЛЕНО: Имя столбца ===
                        cursor.execute('''
                            INSERT INTO chart_data_sources (series_id, data_type, formula)
                            VALUES (?, ?, ?)
                        ''', (series_id, data_type, formula))
                        # ===================================
                        logger.debug(f"    Добавлен источник данных ({data_type}) для серии ID {series_id}")

        connection.commit()
        logger.info(f"Диаграммы для листа ID {sheet_id} сохранены успешно.")
        return True
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении диаграмм для листа ID {sheet_id}: {e}")
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении диаграмм для листа ID {sheet_id}: {e}")
        connection.rollback()
        return False

def load_sheet_charts(connection: sqlite3.Connection, sheet_id: int) -> List[Dict[str, Any]]:
    """
    Загружает диаграммы для указанного листа из структурированных таблиц.
    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа.
    Returns:
        List[Dict[str, Any]]: Список словарей с данными диаграмм.
    """
    charts_data = []
    if not connection:
        logger.error("Получено пустое соединение для загрузки диаграмм.")
        return charts_data
    try:
        cursor = connection.cursor()
        # Загружаем основную информацию о диаграммах
        # === ИСПРАВЛЕНО: Убраны кавычки с "order" и "references" ===
        cursor.execute('''
            SELECT id, type, title, top_left_cell, width, height, style, legend_position, auto_scaling, plot_vis_only
            FROM charts WHERE sheet_id = ?
        ''', (sheet_id,))
        # ========================================================
        charts_rows = cursor.fetchall()
        for chart_row in charts_rows:
            chart_id, chart_type, chart_title, top_left_cell, width, height, chart_style, legend_position, auto_scaling, plot_vis_only = chart_row
            chart_info = {
                "type": chart_type,
                "title": chart_title,
                "top_left_cell": top_left_cell,
                "width": width,
                "height": height,
                "style": chart_style,
                "legend_position": legend_position,
                "auto_scaling": auto_scaling,
                "plot_vis_only": plot_vis_only,
                "axes": [],
                "series": [],
                "data_sources": [] # Будет заполнено позже
            }
            # Загружаем оси для этой диаграммы
            # === ИСПРАВЛЕНО: Имена столбцов соответствуют schema.py ===
            cursor.execute('''
                SELECT axis_type, ax_id, ax_pos, delete_axis, title, num_fmt,
                       major_tick_mark, minor_tick_mark, tick_lbl_pos, crosses, crosses_at,
                       major_unit, minor_unit, min, max, orientation, log_base,
                       major_gridlines, minor_gridlines
                FROM chart_axes WHERE chart_id = ?
            ''', (chart_id,))
            # =========================================================
            axes_rows = cursor.fetchall()
            for axis_row in axes_rows:
                # === ИСПРАВЛЕНО: Имена ключей соответствуют schema.py ===
                axis_info = {
                    "axis_type": axis_row[0], "ax_id": axis_row[1], "ax_pos": axis_row[2],
                    "delete": axis_row[3], "title": axis_row[4], "num_fmt": axis_row[5], # <-- ИСПРАВЛЕНО: delete_axis -> delete, axis_title -> title
                    "major_tick_mark": axis_row[6], "minor_tick_mark": axis_row[7],
                    "tick_lbl_pos": axis_row[8], "crosses": axis_row[9], "crosses_at": axis_row[10],
                    "major_unit": axis_row[11], "minor_unit": axis_row[12], "min": axis_row[13],
                    "max": axis_row[14], "orientation": axis_row[15], "log_base": axis_row[16],
                    "major_gridlines": axis_row[17], "minor_gridlines": axis_row[18]
                }
                # =========================================================
                chart_info["axes"].append(axis_info)
            # Загружаем серии данных для этой диаграммы
            # === ИСПРАВЛЕНО: Убраны кавычки с "order" ===
            cursor.execute('''
                SELECT id, idx, "order", tx, shape, smooth, invert_if_negative
                FROM chart_series WHERE chart_id = ? ORDER BY "order"
            ''', (chart_id,))
            # =================================================
            series_rows = cursor.fetchall()
            series_ids = [] # Собираем ID серий для последующей загрузки источников
            for series_row in series_rows:
                series_id, series_idx, series_order, series_tx, shape, smooth, invert_if_negative = series_row
                series_ids.append(series_id)
                chart_info["series"].append({
                    "idx": series_idx,
                    "order": series_order, # <-- ИСПРАВЛЕНО: "order" -> order
                    "tx": series_tx,
                    "shape": shape,
                    "smooth": smooth,
                    "invert_if_negative": invert_if_negative
                })
            # Загружаем источники данных для всех серий этой диаграммы
            if series_ids:
                placeholders = ','.join('?' * len(series_ids))
                # === ИСПРАВЛЕНО: source_type -> data_type, имена столбцов ===
                cursor.execute(f'''
                    SELECT cds.series_id, cds.data_type, cds.formula, cs.idx
                    FROM chart_data_sources cds
                    JOIN chart_series cs ON cds.series_id = cs.id
                    WHERE cds.series_id IN ({placeholders})
                    ORDER BY cs."order", cds.id
                ''', series_ids)
                # ===============================================
                data_sources_rows = cursor.fetchall()
                for ds_row in data_sources_rows:
                    series_id, data_type, formula, series_idx = ds_row # <-- ИСПРАВЛЕНО: source_type -> data_type
                    chart_info["data_sources"].append({
                        "series_index": series_idx, # Используем idx серии для связи
                        "data_type": data_type, # <-- ИСПРАВЛЕНО
                        "formula": formula
                    })
            charts_data.append(chart_info)
        logger.info(f"Загружено {len(charts_data)} диаграмм для листа ID {sheet_id}.")
        return charts_data
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке диаграмм для листа ID {sheet_id}: {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке диаграмм для листа ID {sheet_id}: {e}")
        return []
