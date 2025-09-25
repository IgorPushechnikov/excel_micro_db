# src/exporter/excel/xlsxwriter_exporter.py
"""
Модуль для экспорта проекта Excel Micro DB в файл Excel (.xlsx) с использованием библиотеки xlsxwriter.
"""

import logging
import json
from pathlib import Path
from typing import Dict, Any, List, Optional, Union

import xlsxwriter # Импортируем xlsxwriter

# Импортируем ProjectDBStorage для взаимодействия с БД
from src.storage.base import ProjectDBStorage

# Импортируем вспомогательные функции для конвертации стилей
from src.exporter.excel.style_handlers.db_style_converter import json_style_to_xlsxwriter_format

# Импортируем ProjectDBStorage для загрузки диаграмм
from src.storage.base import ProjectDBStorage

logger = logging.getLogger(__name__)


def export_project_xlsxwriter(project_db_path: Union[str, Path], output_path: Union[str, Path]) -> bool:
    """
    Основная функция экспорта проекта в Excel файл с помощью xlsxwriter.

    Args:
        project_db_path (Union[str, Path]): Путь к файлу БД проекта (project_data.db).
        output_path (Union[str, Path]): Путь к выходному .xlsx файлу.

    Returns:
        bool: True, если экспорт успешен, иначе False.
    """
    logger.info(f"Начало экспорта проекта в '{output_path}' с использованием xlsxwriter.")
    logger.debug(f"Путь к БД проекта: {project_db_path}")

    project_db_path = Path(project_db_path)
    output_path = Path(output_path)

    if not project_db_path.exists():
        logger.error(f"Файл БД проекта не найден: {project_db_path}")
        return False

    # 1. Подключение к БД проекта
    logger.info("Подключение к БД проекта...")
    try:
        storage = ProjectDBStorage(str(project_db_path))
        if not storage.connect():
            logger.error("Не удалось подключиться к БД проекта.")
            return False
    except Exception as e:
        logger.error(f"Ошибка при создании ProjectDBStorage: {e}")
        return False

    # 2. Создание новой книги xlsxwriter
    logger.info("Создание новой книги Excel с помощью xlsxwriter...")
    try:
        # Создаём директорию для выходного файла, если её нет
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook_options = {
            'strings_to_numbers': True,  # Пытаться конвертировать строки в числа
            'strings_to_formulas': False, # Не пытаться интерпретировать строки как формулы
            'default_date_format': 'dd/mm/yyyy', # Пример формата даты
        }
        workbook = xlsxwriter.Workbook(str(output_path), workbook_options)
    except Exception as e:
        logger.error(f"Ошибка при создании книги xlsxwriter: {e}")
        storage.disconnect() # Закрываем соединение при ошибке
        return False

    success = False
    try:
        # 3. Получение списка листов из БД
        logger.debug("Получение списка листов из БД...")
        sheets_data = storage.load_all_sheets_metadata() # Предполагаем, что в storage есть такой метод
        if not sheets_data:
            logger.warning("В проекте не найдено листов. Создается пустой файл.")
            workbook.add_worksheet("EmptySheet")
        else:
            logger.info(f"Найдено {len(sheets_data)} листов для экспорта.")
            # 4. Итерация по листам и их экспорт
            for sheet_info in sheets_data:
                sheet_id = sheet_info['sheet_id']
                sheet_name = sheet_info['name']
                logger.info(f"Экспорт листа: '{sheet_name}' (ID: {sheet_id})")

                # 4a. Создание листа в xlsxwriter
                worksheet = workbook.add_worksheet(sheet_name)

                # 4b. Загрузка данных для листа
                # Предполагаем, что storage предоставляет методы для загрузки данных
                raw_data = storage.load_sheet_raw_data(sheet_name) # Возвращает список {'cell_address': ..., 'value': ...}
                formulas = storage.load_sheet_formulas(sheet_id) # Возвращает список {'cell_address': ..., 'formula': ...}
                styles = storage.load_sheet_styles(sheet_id) # Возвращает список {'range_address': ..., 'style_attributes': ...}
                merged_cells = storage.load_sheet_merged_cells(sheet_id) # Возвращает список ['A1:B2', ...]
                logger.debug(f"[ЭКСПОРТ] Загружены объединённые ячейки для листа '{sheet_name}' (ID: {sheet_id}): {merged_cells}")

                # 4c. Подготовка стилей (создание карты форматов)
                cell_format_map = build_cell_format_map(workbook, styles)

                # 4d. Запись данных и формул с применением стилей
                written_cells = _write_data_and_formulas(worksheet, raw_data, formulas, cell_format_map)

                # 4e. Применение стилей к пустым ячейкам, у которых есть стиль в cell_format_map
                for (r, c), cell_format in cell_format_map.items():
                    if (r, c) not in written_cells:
                        # Используем write_blank для установки формата на пустую ячейку
                        worksheet.write_blank(r, c, None, cell_format)
                        logger.debug(f"Применён стиль к пустой ячейке ({r}, {c})")

                # 4f. Применение объединенных ячеек
                logger.debug(f"[ЭКСПОРТ] Перед вызовом _apply_merged_cells для листа '{sheet_name}' с данными: {merged_cells}")
                _apply_merged_cells(worksheet, merged_cells)

                # 4g. Экспорт диаграмм
                logger.debug(f"[ЭКСПОРТ] Перед вызовом _export_charts_for_sheet для листа '{sheet_name}' (ID: {sheet_id})")
                _export_charts_for_sheet(workbook, worksheet, sheet_id, project_db_path)

                # 4h. (Опционально) Обработка других элементов (диаграмм, изображений и т.д.)
                # ...

        # 5. Закрытие книги (сохранение файла)
        logger.info("Закрытие книги и сохранение файла...")
        workbook.close()
        logger.info(f"Файл успешно сохранен: {output_path}")
        success = True

    except Exception as e:
        logger.error(f"Критическая ошибка при экспорте проекта: {e}", exc_info=True)
        # workbook.close() вызывается автоматически при выходе из блока try/except,
        # если он был открыт, но xlsxwriter может не сохранить файл при ошибке.
        # Важно, чтобы storage.disconnect() вызывался в finally.

    finally:
        # 6. Закрытие соединения с БД
        logger.info("Закрытие соединения с БД проекта.")
        storage.disconnect()

    return success


def _write_data_and_formulas(worksheet, raw_data: List[Dict[str, Any]], formulas: List[Dict[str, Any]], cell_format_map: Dict[tuple[int, int], Any]) -> set[tuple[int, int]]:
    """
    Записывает данные и формулы на лист xlsxwriter, применяя стили из cell_format_map.
    Возвращает множество координат (row, col), в которые что-то было записано.

    Args:
        worksheet: Объект листа xlsxwriter.
        raw_data (List[Dict[str, Any]]): Список данных.
        formulas (List[Dict[str, Any]]): Список формул.
        cell_format_map (Dict[tuple[int, int], Any]): Словарь сопоставления (row, col) -> xlsxwriter.format.

    Returns:
        set[tuple[int, int]]: Множество координат (row, col), в которые были записаны данные или формулы.
    """
    logger.debug(f"Запись {len(raw_data)} записей данных и {len(formulas)} формул на лист с применением стилей.")
    written_cells = set()
    # Запись "сырых" данных
    for item in raw_data:
        address = item['cell_address'] # e.g., 'A1'
        value = item['value']
        # xlsxwriter требует номера строки/столбца, преобразуем адрес
        try:
            row, col = _xl_cell_to_row_col(address)
            # Проверяем, есть ли формат для этой ячейки
            cell_format = cell_format_map.get((row, col))
            worksheet.write(row, col, value, cell_format)
            written_cells.add((row, col))
        except Exception as e:
            logger.warning(f"Не удалось записать данные в ячейку {address}: {e}")

    # Запись формул
    for item in formulas:
        address = item['cell_address']
        formula = item['formula']
        try:
            row, col = _xl_cell_to_row_col(address)
            # Для формул xlsxwriter ожидает строку без '='
            formula_clean = formula[1:] if formula.startswith('=') else formula
            # Проверяем, есть ли формат для этой ячейки
            cell_format = cell_format_map.get((row, col))
            worksheet.write_formula(row, col, formula_clean, cell_format)
            written_cells.add((row, col))
        except Exception as e:
            logger.warning(f"Не удалось записать формулу в ячейку {address}: {e}")
    
    return written_cells


def build_cell_format_map(workbook, styles: List[Dict[str, Any]]) -> Dict[tuple[int, int], Any]:
    """
    Создает словарь, сопоставляющий координаты ячеек (row, col) с форматами xlsxwriter.
    Это позволяет применять стили одновременно с записью данных/формул.

    Args:
        workbook: Объект книги xlsxwriter (для создания форматов).
        styles (List[Dict[str, Any]]): Список стилей из БД.

    Returns:
        Dict[tuple[int, int], Any]: Словарь, где ключ - (row, col), значение - объект формата xlsxwriter.
    """
    logger.debug(f"Создание карты форматов ячеек для {len(styles)} стилей.")
    cell_format_map: Dict[tuple[int, int], Any] = {}

    for style_item in styles:
        range_addr = style_item['range_address'] # e.g., 'A1:B10'
        style_json_str = style_item['style_attributes']

        try:
            # 1. Конвертируем JSON-стиль в формат xlsxwriter
            xlsxwriter_format_dict = json_style_to_xlsxwriter_format(style_json_str)
            if not xlsxwriter_format_dict:
                logger.debug(f"Для стиля {range_addr} не определено атрибутов для xlsxwriter, пропуск.")
                continue

            # 2. Создаём формат xlsxwriter
            cell_format = workbook.add_format(xlsxwriter_format_dict)

            # 3. Заполняем карту форматов для каждой ячейки в диапазоне
            row_start, col_start, row_end, col_end = _xl_range_to_coords(range_addr)
            
            for r in range(row_start, row_end + 1):
                for c in range(col_start, col_end + 1):
                    # Если ячейка уже имеет формат, xlsxwriter использует первый применённый формат.
                    # В реальных сценариях диапазоны могут пересекаться, и нужно решать, какой стиль приоритетнее.
                    # Для MVP/простоты принимаем первый встреченный стиль для ячейки.
                    if (r, c) not in cell_format_map:
                        cell_format_map[(r, c)] = cell_format
                    else:
                        # Логируем, если стили пересекаются, чтобы было видно в отладке
                        logger.debug(f"Ячейка ({r}, {c}) уже имеет формат. Второй стиль из диапазона {range_addr} игнорируется.")

        except json.JSONDecodeError as je:
            logger.error(f"Ошибка разбора JSON стиля для диапазона {range_addr}: {je}")
        except Exception as e:
            logger.error(f"Ошибка при создании карты форматов для диапазона {range_addr}: {e}", exc_info=True)
    
    logger.debug(f"Создана карта форматов для {len(cell_format_map)} ячеек.")
    return cell_format_map


def _apply_styles(workbook, worksheet, styles: List[Dict[str, Any]]):
    """
    Применяет стили к диапазонам на листе xlsxwriter.
    В текущей реализации создает карту форматов и передает её в _write_data_and_formulas.

    Args:
        workbook: Объект книги xlsxwriter.
        worksheet: Объект листа xlsxwriter.
        styles (List[Dict[str, Any]]): Список стилей.
    """
    logger.debug(f"Применение {len(styles)} стилей к листу через карту форматов.")
    cell_format_map = build_cell_format_map(workbook, styles)
    # Стили теперь готовы к применению при записи данных/формул
    # _apply_styles больше не записывает напрямую, только подготавливает данные
    logger.debug(f"Карта форматов создана. Передача в _write_data_and_formulas.")


def _xl_cell_to_row_col(cell: str) -> tuple[int, int]:
    """
    Преобразует адрес ячейки Excel (e.g., 'A1') в индексы строки и столбца (0-based).
    """
    logger.debug(f"[КООРД] Преобразование ячейки '{cell}' в (row, col) (0-based).")
    from openpyxl.utils import coordinate_to_tuple
    # Используем вспомогательную функцию из openpyxl, она надежна.
    # row, col = coordinate_to_tuple(cell) # row, col are 1-based
    # return row - 1, col - 1 # Convert to 0-based
    # Или реализуем вручную, чтобы не зависеть от openpyxl в этом модуле.
    col_str = ""
    row_str = ""
    for char in cell:
        if char.isalpha():
            col_str += char.upper()
        elif char.isdigit():
            row_str += char

    if not col_str or not row_str:
        error_msg = f"Неверный формат адреса ячейки: {cell}"
        logger.error(f"[КООРД] {error_msg}")
        raise ValueError(error_msg)

    row = int(row_str) - 1 # 0-based
    col = 0
    for c in col_str:
        col = col * 26 + (ord(c) - ord('A') + 1)
    col -= 1 # 0-based
    result = (row, col)
    logger.debug(f"[КООРД] Результат для '{cell}': {result}")
    return result


def _xl_range_to_coords(range_str: str) -> tuple[int, int, int, int]:
    """
    Преобразует диапазон Excel (e.g., 'A1:B10') в координаты (row_start, col_start, row_end, col_end) (0-based).
    """
    logger.debug(f"[КООРД] Преобразование диапазона '{range_str}' в координаты.")
    if ':' not in range_str:
        # Это одиночная ячейка
        logger.debug(f"[КООРД] Диапазон '{range_str}' - это одиночная ячейка.")
        r, c = _xl_cell_to_row_col(range_str)
        coords = (r, c, r, c)
        logger.debug(f"[КООРД] Результат для '{range_str}': {coords}")
        return coords

    start_cell, end_cell = range_str.split(':', 1)
    logger.debug(f"[КООРД] Разделение диапазона на '{start_cell}' и '{end_cell}'.")
    row_start, col_start = _xl_cell_to_row_col(start_cell)
    row_end, col_end = _xl_cell_to_row_col(end_cell)
    coords = (row_start, col_start, row_end, col_end)
    logger.debug(f"[КООРД] Результат для диапазона '{range_str}': {coords}")
    return coords


def _export_charts_for_sheet(workbook, worksheet, sheet_id: int, project_db_path: Union[str, Path]):
    """
    Экспортирует диаграммы для указанного листа, используя xlsxwriter.

    Args:
        workbook: Объект книги xlsxwriter.
        worksheet: Объект листа xlsxwriter.
        sheet_id (int): ID листа в БД проекта.
        project_db_path (Union[str, Path]): Путь к файлу БД проекта.
    """
    logger.info(f"[ДИАГРАММА] Начало экспорта диаграмм для листа ID {sheet_id}.")
    
    # 1. Подключаемся к БД проекта для загрузки диаграмм
    storage = ProjectDBStorage(str(project_db_path))
    if not storage.connect():
        logger.error(f"[ДИАГРАММА] Не удалось подключиться к БД проекта для загрузки диаграмм листа ID {sheet_id}.")
        return
    
    try:
        # 2. Загружаем диаграммы из БД
        charts_data = storage.load_sheet_charts(sheet_id)
        logger.debug(f"[ДИАГРАММА] Загружено {len(charts_data)} диаграмм для листа ID {sheet_id}.")
        
        if not charts_data:
            logger.info(f"[ДИАГРАММА] Нет диаграмм для экспорта на листе ID {sheet_id}.")
            storage.disconnect()
            return
        
        # 3. Итерация по диаграммам и их экспорт
        for chart_entry in charts_data:
            chart_data_json_str = chart_entry.get('chart_data')
            if not chart_data_json_str:
                logger.warning(f"[ДИАГРАММА] Найдена запись диаграммы без данных. Пропущена.")
                continue
            
            try:
                # Десериализуем JSON-строку в словарь
                chart_data = json.loads(chart_data_json_str)
                logger.debug(f"[ДИАГРАММА] Обработка диаграммы типа: {chart_data.get('type', 'Unknown')}.")
                
                # 4. Создаем объект диаграммы xlsxwriter
                # Сопоставляем тип диаграммы из openpyxl с типом xlsxwriter
                chart_type_map = {
                    'BarChart': 'column', # xlsxwriter использует 'column' для столбчатых диаграмм
                    'LineChart': 'line',
                    'PieChart': 'pie',
                    'PieChart3D': 'pie', # xlsxwriter не поддерживает 3D pie напрямую через тип, но можно установить 3D опции
                    # ... другие типы
                }
                xlsxwriter_chart_type = chart_type_map.get(chart_data['type'], 'column') # Дефолт 'column'
                
                chart_options = {'type': xlsxwriter_chart_type}
                # Для 3D диаграмм добавляем опции
                if chart_data['type'] == 'PieChart3D':
                    chart_options['subtype'] = '3d'
                
                chart = workbook.add_chart(chart_options)
                logger.debug(f"[ДИАГРАММА] Создан объект xlsxwriter.Chart типа: {xlsxwriter_chart_type}.")
                
                # 5. Настраиваем данные диаграммы (series)
                series_list = chart_data.get('series', [])
                for series in series_list:
                    val_range = series.get('val_range')
                    cat_range = series.get('cat_range')
                    name = series.get('name')
                    
                    if not val_range:
                        logger.warning(f"[ДИАГРАММА] Серия диаграммы не содержит диапазон значений (val_range). Пропущена.")
                        continue
                    
                    series_options = {'values': val_range}
                    if cat_range:
                        series_options['categories'] = cat_range
                    if name:
                        series_options['name'] = name
                    
                    chart.add_series(series_options)
                    logger.debug(f"[ДИАГРАММА] Добавлена серия: values={val_range}, categories={cat_range}, name={name}")
                
                # 6. Настраиваем заголовок диаграммы
                title = chart_data.get('title')
                title_ref = chart_data.get('title_ref')
                if title is not None: # Даже если title == ""
                    chart.set_title({'name': title})
                    logger.debug(f"[ДИАГРАММА] Установлен заголовок: '{title}'")
                elif title_ref:
                    chart.set_title({'name': title_ref}) # xlsxwriter может интерпретировать ссылку
                    logger.debug(f"[ДИАГРАММА] Установлена ссылка на заголовок: '{title_ref}'")
                
                # 7. Настраиваем размеры диаграммы
                width_emu = chart_data.get('width_emu')
                height_emu = chart_data.get('height_emu')
                if width_emu is not None and height_emu is not None:
                    # xlsxwriter.set_size ожидает размеры в пикселях
                    # Конвертируем EMU в пиксели (приблизительно, 914400 EMU = 1 дюйм, 96 DPI)
                    # pixels = emu / 914400 * 96
                    # Упрощаем: 1 EMU = 1/914400 дюйма, 1 дюйм = 96 пикселей => 1 EMU = 96/914400 пикселей
                    # 96/914400 = 1/9525
                    width_px = int(width_emu / 9525)
                    height_px = int(height_emu / 9525)
                    chart.set_size({'width': width_px, 'height': height_px})
                    logger.debug(f"[ДИАГРАММА] Установлен размер: {width_px}x{height_px} пикселей (из EMU {width_emu}x{height_emu})")
                
                # 8. Настраиваем позицию и вставляем диаграмму на лист
                position_info = chart_data.get('position')
                if position_info:
                from_col = position_info.get('from_col')
                from_row = position_info.get('from_row')
                from_col_offset_emu = position_info.get('from_col_offset')
                from_row_offset_emu = position_info.get('from_row_offset')

                if from_col is not None and from_row is not None:
                # xlsxwriter.insert_chart может принимать смещения в опциях
                # Конвертируем EMU в пиксели для смещений
                x_offset_px = int(from_col_offset_emu / 9525) if from_col_offset_emu is not None else 0
                y_offset_px = int(from_row_offset_emu / 9525) if from_row_offset_emu is not None else 0

                # Вставляем с опциями смещения
                insert_options = {'x_offset': x_offset_px, 'y_offset': y_offset_px}
                    worksheet.insert_chart(from_row, from_col, chart, insert_options)
                logger.debug(f"[ДИАГРАММА] Диаграмма вставлена в ({from_row}, {from_col}) с опциями: {insert_options}")
                    else:
                    logger.warning(f"[ДИАГРАММА] Неполные данные позиции: from_col={from_col}, from_row={from_row}")
                        # Вставляем без смещения, если данные неполные
                        worksheet.insert_chart(from_row, from_col, chart)
                else:
                    logger.warning(f"[ДИАГРАММА] У диаграммы нет информации о позиции. Будет размещена по умолчанию.")
                    # Если позиция неизвестна, вставляем в ячейку A1 (0, 0)
                    worksheet.insert_chart(0, 0, chart) # Вставляем в A1
                
                logger.info(f"[ДИАГРАММА] Диаграмма типа '{chart_data['type']}' успешно экспортирована на лист ID {sheet_id}.")
                
            except json.JSONDecodeError as je:
                logger.error(f"[ДИАГРАММА] Ошибка разбора JSON данных диаграммы: {je}")
            except Exception as e_inner:
                logger.error(f"[ДИАГРАММА] Ошибка при обработке одной из диаграмм для листа ID {sheet_id}: {e_inner}", exc_info=True)
        
        # 10. Закрываем соединение с БД
        storage.disconnect()
        logger.info(f"[ДИАГРАММА] Экспорт диаграмм для листа ID {sheet_id} завершен.")
        
    except Exception as e_outer:
        logger.error(f"[ДИАГРАММА] Критическая ошибка при экспорте диаграмм для листа ID {sheet_id}: {e_outer}", exc_info=True)
        # Пытаемся закрыть соединение, если оно было открыто
        if storage:
            storage.disconnect()

def _apply_merged_cells(worksheet, merged_ranges: List[str]):
    """
    Применяет объединения ячеек к листу xlsxwriter.

    Args:
        worksheet: Объект листа xlsxwriter.
        merged_ranges (List[str]): Список строковых адресов диапазонов (например, ['A1:B2', 'C3:D5']).
    """
    logger.debug(f"[ОБЪЕДИНЕНИЕ] Начало применения объединений. Получено {len(merged_ranges)} диапазонов: {merged_ranges}")
    if not merged_ranges:
        logger.debug("[ОБЪЕДИНЕНИЕ] Список диапазонов пуст. Нечего применять.")
        return
    
    applied_count = 0
    for range_addr in merged_ranges:
        try:
            logger.debug(f"[ОБЪЕДИНЕНИЕ] Обработка диапазона: '{range_addr}'")
            if not range_addr or ":" not in range_addr:
                 logger.warning(f"[ОБЪЕДИНЕНИЕ] Неверный формат диапазона объединения: '{range_addr}'. Пропущен.")
                 continue

            # xlsxwriter.merge_range требует (first_row, first_col, last_row, last_col)
            first_row, first_col, last_row, last_col = _xl_range_to_coords(range_addr)
            logger.debug(f"[ОБЪЕДИНЕНИЕ] Координаты диапазона '{range_addr}': ({first_row}, {first_col}) -> ({last_row}, {last_col})")
            
            # merge_range также требует значение и формат. Передаем None и None.
            # Если нужно заполнить объединенную ячейку данными или стилем, логика усложняется.
            # Пока просто объединяем.
            worksheet.merge_range(first_row, first_col, last_row, last_col, None)
            logger.info(f"[ОБЪЕДИНЕНИЕ] Успешно объединен диапазон: {range_addr}")
            applied_count += 1
            
        except ValueError as ve: # Ошибка от _xl_range_to_coords
            logger.error(f"[ОБЪЕДИНЕНИЕ] Ошибка преобразования координат диапазона '{range_addr}': {ve}")
        except Exception as e:
            logger.error(f"[ОБЪЕДИНЕНИЕ] Критическая ошибка при объединении диапазона '{range_addr}': {e}", exc_info=True)
    
    logger.info(f"[ОБЪЕДИНЕНИЕ] Завершено. Успешно применено {applied_count}/{len(merged_ranges)} объединений.")
