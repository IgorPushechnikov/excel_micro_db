# src/analyzer/logic_documentation.py

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
# --- ИСПРАВЛЕНИЕ: Импортируем Cell напрямую с псевдонимом ---
# Это должно помочь Pylance понять тип, совместимый с _CellOrMergedCell
from openpyxl.cell.cell import Cell as OpenPyxlCell
# Для аннотаций типов
from typing import Dict, Any, List, Optional
import logging
import json

# Импортируем logger из utils
from src.utils.logger import get_logger

logger = get_logger(__name__)

# --- Вспомогательные функции для сериализации сложных объектов ---

# --- ИСПРАВЛЕНИЕ: Используем новый синтаксис объединения типов ---
# В Python 3.10+ можно использовать X | Y для объединения типов
# iter_rows() возвращает _CellOrMergedCell, который должен быть совместим с OpenPyxlCell.
# В новых версиях Python Pylance должен принять эту аннотацию.
def _serialize_style(cell: OpenPyxlCell | Any) -> Dict[str, Any]:
    """
    Сериализует атрибуты стиля ячейки openpyxl в словарь.
    Этот словарь будет сериализован в JSON в storage/styles.py.
    Структура должна соответствовать ожиданиям _convert_style_to_xlsxwriter_format
    в excel_exporter.py.
    """
    # --- ИСПРАВЛЕНИЕ: Pylance теперь должен понимать, что у cell (OpenPyxlCell | Any) есть атрибуты font, fill и т.д.
    # Проверки hasattr остаются для безопасности во время выполнения.
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
    try:
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
                 if hasattr(chart_obj.title.tx, 'rich') and chart_obj.title.tx.rich:
                     try:
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

        logger.debug(f"Сериализована диаграмма типа {chart_data.get('type', 'Unknown')}")
        return chart_data

    except Exception as e:
        logger.error(f"Ошибка при сериализации диаграммы: {e}", exc_info=True)
        return {}

# --- Основная функция анализа ---

def analyze_excel_file(file_path: str) -> Dict[str, Any]:
    """
    Основная функция для анализа Excel-файла.
    Извлекает структуру, данные, формулы, стили, диаграммы и другую информацию.
    Возвращает словарь с результатами анализа, готовый для передачи в storage.

    Args:
        file_path (str): Путь к анализируемому .xlsx файлу.

    Returns:
        Dict[str, Any]: Словарь с результатами анализа.
    """
    logger.info(f"Начало анализа Excel-файла: {file_path}")

    try:
        # Открываем книгу Excel
        workbook = openpyxl.load_workbook(file_path, data_only=False) # data_only=False для получения формул
        logger.debug(f"Книга '{file_path}' успешно открыта.")

        analysis_results = {
            "project_name": file_path.split("/")[-1].split(".")[0], # Имя проекта из имени файла
            "file_path": file_path,
            "sheets": []
        }

        # Итерируемся по всем листам в книге
        for sheet_name in workbook.sheetnames:
            logger.info(f"Анализ листа: {sheet_name}")
            sheet: Worksheet = workbook[sheet_name]

            sheet_data = {
                "name": sheet_name,
                "max_row": sheet.max_row,
                "max_column": sheet.max_column,
                "raw_data": [],
                "formulas": [],
                "styles": [], # Будет содержать {'range_address': str, 'style_attributes': str (JSON)}
                "charts": [], # Будет содержать {'chart_data': dict или str}
                "merged_cells": [] # Список строк адресов объединенных ячеек
            }

            # --- 1. Извлечение "сырых данных" ---
            logger.debug(f"Извлечение сырых данных с листа '{sheet_name}'...")
            for row in sheet.iter_rows(values_only=False): # values_only=False, чтобы получить объекты Cell
                for cell in row:
                    if cell.value is not None or cell.data_type == 'f': # Сохраняем и данные, и формулы
                        data_item = {
                            "cell_address": cell.coordinate,
                            "value": cell.value,
                            # "value_type": type(cell.value).__name__ # Может быть полезно
                        }
                        sheet_data["raw_data"].append(data_item)

            # --- 2. Извлечение формул ---
            logger.debug(f"Извлечение формул с листа '{sheet_name}'...")
            for row in sheet.iter_rows(values_only=False):
                for cell in row:
                    # Проверим, является ли значение строкой, начинающейся с '='
                    # Или используем data_type для надежности
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                         sheet_data["formulas"].append({
                             "cell_address": cell.coordinate,
                             "formula": cell.value # Сохраняем формулу как есть, включая '='
                         })
                    # ВАЖНО: openpyxl также имеет cell.formula, но это может быть не то же самое.
                    # Уточнение: cell.value при data_only=False содержит формулу.

            # --- 3. Извлечение стилей ---
            logger.debug(f"Извлечение и группировка стилей с листа '{sheet_name}'...")
            # Простой подход: для каждой ячейки сохраняем её стиль.
            # Более сложный (эффективный): группировать ячейки с одинаковыми стилями в диапазоны.
            # Пока используем простой подход для MVP.

            # Словарь для хранения уникальных стилей и их диапазонов
            style_ranges_map: Dict[str, List[str]] = {} # ключ - сериализованный стиль, значение - список адресов

            for row in sheet.iter_rows(values_only=False):
                for cell in row:
                    # Проверяем, есть ли у ячейки не-дефолтный стиль
                    # Это можно сделать, сравнивая с дефолтным стилем, но проще сериализовать всегда
                    # и потом в storage/styles.py решать, нужно ли его сохранять.

                    # --- ИСПРАВЛЕНИЕ: Передаем cell в _serialize_style, аннотированную как OpenPyxlCell | Any ---
                    # iter_rows возвращает _CellOrMergedCell. Pylance должен принять OpenPyxlCell | Any.
                    style_dict = _serialize_style(cell)
                    # --- КОНЕЦ ИСПРАВЛЕНИЯ ---
                    if style_dict: # Если стиль не пустой
                        style_json = json.dumps(style_dict, sort_keys=True) # Сериализуем в JSON и сортируем ключи для надежного хеширования
                        coord = cell.coordinate

                        if style_json in style_ranges_map:
                            style_ranges_map[style_json].append(coord)
                        else:
                            style_ranges_map[style_json] = [coord]

            # Преобразуем карту стилей в формат, ожидаемый storage
            for style_json, cell_addresses in style_ranges_map.items():
                # Для упрощения, будем создавать отдельную запись для каждой ячейки
                # В будущем можно реализовать группировку в диапазоны (A1:A10, B1:D1 и т.д.)
                # Это требует сложной логики группировки.
                for address in cell_addresses:
                     sheet_data["styles"].append({
                         "range_address": address, # Пока каждая ячейка отдельно
                         "style_attributes": style_json # Строка JSON
                     })
                # TODO: Реализовать группировку адресов в диапазоны
                # sheet_data["styles"].append({
                #     "range_address": _group_addresses(cell_addresses), # Функция группировки
                #     "style_attributes": style_json # Строка JSON
                # })

            # --- 4. Извлечение диаграмм ---
            logger.debug(f"Извлечение диаграмм с листа '{sheet_name}'...")
            # Диаграммы находятся в sheet._charts
            # --- ИСПРАВЛЕНИЕ: Оборачиваем доступ к _charts в try/except и добавляем # type: ignore ---
            try:
                # Атрибут _charts является внутренним. Доступ к нему может быть ненадежным.
                # Используем # type: ignore[attr-defined] для подавления предупреждения Pylance.
                charts_list = sheet._charts # type: ignore[attr-defined]
                for chart_obj in charts_list:
                     chart_data = _serialize_chart(chart_obj)
                     if chart_data:
                         # storage ожидает 'chart_data' как сериализованный объект
                         sheet_data["charts"].append({
                             "chart_data": chart_data # Это будет словарь, storage должен его сериализовать при сохранении
                             # Если нужно сохранить как JSON сразу:
                             # "chart_data": json.dumps(chart_data, ensure_ascii=False)
                         })
            except AttributeError as ae:
                logger.warning(f"Не удалось получить доступ к диаграммам листа '{sheet_name}' через _charts: {ae}")
            except Exception as e:
                logger.error(f"Ошибка при извлечении диаграмм с листа '{sheet_name}': {e}", exc_info=True)
            # --- КОНЕЦ ИСПРАВЛЕНИЯ ---


            # --- 5. Извлечение объединенных ячеек ---
            logger.debug(f"Извлечение объединенных ячеек с листа '{sheet_name}'...")
            for merged_cell_range in sheet.merged_cells.ranges:
                # merged_cell_range это openpyxl.utils.cell_range.CellRange
                sheet_data["merged_cells"].append(str(merged_cell_range)) # Преобразуем в строку адреса диапазона

            # Добавляем данные листа в результаты анализа
            analysis_results["sheets"].append(sheet_data)

        logger.info(f"Анализ Excel-файла '{file_path}' завершен.")
        return analysis_results

    except Exception as e:
        logger.error(f"Ошибка при анализе Excel-файла '{file_path}': {e}", exc_info=True)
        # Возвращаем пустой словарь или поднимаем исключение
        # В реальном приложении лучше поднимать пользовательское исключение
        raise # Повторно поднимаем исключение для обработки выше

# --- Функция для группировки адресов ячеек в диапазоны (заглушка) ---
# def _group_addresses(addresses: List[str]) -> str:
#     """
#     Группирует список адресов ячеек в строку диапазонов.
#     Например: ['A1', 'A2', 'A3', 'B1'] -> 'A1:A3 B1'
#     Это сложная задача, требующая алгоритмов.
#     """
#     # TODO: Реализовать алгоритм группировки
#     # Пока возвращаем объединение через пробел
#     return " ".join(addresses)

# Пример использования (если файл запускается напрямую)
if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Использование: python logic_documentation.py <excel_file_path>")
        sys.exit(1)

    file_path = sys.argv[1]
    try:
        results = analyze_excel_file(file_path)
        print(f"Анализ завершен. Результаты для {results['project_name']}:")
        print(f"  - Листов: {len(results['sheets'])}")
        for sheet in results['sheets']:
            print(f"    - Лист '{sheet['name']}':")
            print(f"      - Ячеек с данными: {len(sheet['raw_data'])}")
            print(f"      - Формул: {len(sheet['formulas'])}")
            print(f"      - Стилей: {len(sheet['styles'])}")
            print(f"      - Диаграмм: {len(sheet['charts'])}")
            print(f"      - Объединенных ячеек: {len(sheet['merged_cells'])}")
    except Exception as e:
        print(f"Ошибка при анализе: {e}")
        sys.exit(1)