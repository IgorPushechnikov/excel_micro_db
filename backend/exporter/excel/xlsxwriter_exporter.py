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
from storage.base import ProjectDBStorage

# Импортируем вспомогательные функции для конвертации стилей
from exporter.excel.style_handlers.db_style_converter import json_style_to_xlsxwriter_format

# Импортируем ProjectDBStorage для загрузки диаграмм
from storage.base import ProjectDBStorage

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
                # xlsxwriter не поддерживает subtype '3d' для 'pie'. Опции 3D диаграмм зависят от стиля.
                # subtype строка убирается.
                
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
                        # Если name - это ссылка на ячейку (начинается с '='), извлекаем значение
                        if isinstance(name, str) and name.startswith('='):
                            name_value = _extract_single_value_from_ref(project_db_path, name)
                            if name_value is not None:
                                series_options['name'] = name_value
                            else:
                                logger.warning(f"[ДИАГРАММА] Не удалось извлечь значение для имени ряда из ссылки: {name}")
                        else:
                            series_options['name'] = name
                    
                    chart.add_series(series_options)
                    logger.debug(f"[ДИАГРАММА] Добавлена серия: values={val_range}, categories={cat_range}, name={name}")
                
                # 6. Настраиваем заголовок диаграммы
                title = chart_data.get('title')
                title_ref = chart_data.get('title_ref')
                chart_title_to_set = None
                if title is not None: # Даже если title == ""
                    chart_title_to_set = title
                    logger.debug(f"[ДИАГРАММА] Используется текстовый заголовок: '{title}'")
                elif title_ref:
                    # Пытаемся обработать ссылку на заголовок
                    parsed_ref = _parse_title_ref(title_ref)
                    if parsed_ref:
                        if parsed_ref['is_single_cell']:
                            # Если ссылка на одну ячейку, передаем её как есть
                            chart_title_to_set = title_ref
                            logger.debug(f"[ДИАГРАММА] Используется ссылка на одну ячейку как заголовок: '{title_ref}'")
                        else:
                            # Если ссылка на диапазон, пытаемся извлечь значения
                            # TODO: Реализовать извлечение значений из диапазона.
                            # Пока устанавливаем саму ссылку как заголовок, надеясь, что xlsxwriter сможет её обработать.
                            # Если нет, можно установить фиктивный заголовок или оставить пустым.
                            chart_title_to_set = title_ref # Или "Заголовок из " + title_ref
                            logger.debug(f"[ДИАГРАММА] Используется ссылка на диапазон как заголовок (временное решение): '{title_ref}'")
                    else:
                        logger.warning(f"[ДИАГРАММА] Неверный формат title_ref '{title_ref}'. Заголовок не установлен.")
                
                if chart_title_to_set is not None:
                    chart.set_title({'name': chart_title_to_set})
                    logger.debug(f"[ДИАГРАММА] Установлен окончательный заголовок диаграммы: '{chart_title_to_set}'")
                else:
                    logger.debug(f"[ДИАГРАММА] Заголовок диаграммы не будет установлен.")
                
                # --- НОВОЕ: Устанавливаем стиль диаграммы, если она была 3D ---
                if chart_data['type'] == 'PieChart3D':
                    # Попробуем стиль 10, который обычно соответствует 3D Pie Chart в Excel
                    # Стили 10-12 часто бывают 3D для Pie/Doughnut
                    chart.set_style(10) # или 11, 12 - можно протестировать
                    logger.debug(f"[ДИАГРАММА] Установлен 3D стиль (10) для PieChart3D.")
                # --- КОНЕЦ НОВОГО КОДА ---

                # --- НОВОЕ (улучшенное): Настраиваем легенду диаграммы ---
                # Проверяем, есть ли информация о легенде в данных из БД
                legend_info = chart_data.get('legend')
                legend_options = {}

                if legend_info:
                    # Используем настройки легенды из БД
                    # Позиция
                    db_position = legend_info.get('position')
                    # xlsxwriter использует немного другие названия позиций, иногда
                    # Нужно сопоставить. Например, openpyxl 'b' -> xlsxwriter 'bottom'
                    position_mapping = {
                        'b': 'bottom',
                        't': 'top',
                        'l': 'left',
                        'r': 'right',
                        'tr': 'top_right',
                        # Добавить другие сопоставления при необходимости
                    }
                    xlsxwriter_position = position_mapping.get(db_position, db_position) # Используем сопоставление или оригинальное значение
                    if xlsxwriter_position:
                        legend_options['position'] = xlsxwriter_position
                
                # Если из БД не пришло никаких опций, xlsxwriter может показать легенду по умолчанию.
                # Но чтобы быть уверенным, что легенда всегда отображается (если она была в оригинале),
                # и управлять её поведением, мы всегда вызываем set_legend.
                # Если legend_options пуст, будут применены настройки по умолчанию xlsxwriter.
                chart.set_legend(legend_options)
                logger.debug(f"[ДИАГРАММА] Установлена легенда диаграммы с опциями: {legend_options if legend_options else '(по умолчанию)'}")
                # --- КОНЕЦ НОВОГО КОДА ---
                
                # 7. Настраиваем позицию и размер диаграммы, используя ячейки
                # Используем from_row/col и to_row/col для определения размера в ячейках
                position_info = chart_data.get('position')
                if position_info:
                    from_col = position_info.get('from_col')
                    from_row = position_info.get('from_row')
                    to_col = position_info.get('to_col')
                    to_row = position_info.get('to_row')

                    if from_col is not None and from_row is not None:
                        # Вычисляем ширину и высоту в количестве ячеек
                        width_cells = (to_col - from_col + 1) if to_col is not None else 8 # значение по умолчанию
                        height_cells = (to_row - from_row + 1) if to_row is not None else 16 # значение по умолчанию

                        # Вставляем диаграмму с указанием размера в ячейках
                        # xlsxwriter.insert_chart не принимает смещения (offsets) в EMU, только в пикселях
                        # Используем масштабирование через x_scale и y_scale
                        # Эти значения определяют, сколько "ячеек" будет занимать диаграмма
                        # Это приближенное, но более надежное решение, чем конвертация EMU в пиксели
                        chart_options_for_insert = {
                            'x_scale': width_cells / 8.0,  # Масштаб по X (8 - стандартная ширина ячейки в условных единицах)
                            'y_scale': height_cells / 16.0 # Масштаб по Y (16 - стандартная высота ячейки в условных единицах)
                        }
                        worksheet.insert_chart(from_row, from_col, chart, chart_options_for_insert)
                        logger.debug(f"[ДИАГРАММА] Диаграмма вставлена в ({from_row}, {from_col}) с масштабом {width_cells}x{height_cells} ячеек (x_scale={chart_options_for_insert['x_scale']}, y_scale={chart_options_for_insert['y_scale']}).")
                    else:
                        logger.warning(f"[ДИАГРАММА] Неполные данные позиции: from_col={from_col}, from_row={from_row}")
                        # Вставляем в A1 (0, 0), если данные неполные
                        worksheet.insert_chart(0, 0, chart)
                        logger.debug(f"[ДИАГРАММА] Диаграмма вставлена в (0, 0) по умолчанию.")
                else:
                    logger.warning(f"[ДИАГРАММА] У диаграммы нет информации о позиции. Вставляется в (0, 0) по умолчанию.")
                    # Если позиция неизвестна, вставляем в ячейку A1 (0, 0)
                    worksheet.insert_chart(0, 0, chart) # Вставляем в A1

                # УБРАНО: chart.set_size больше не используется, так как размер задается через масштаб при вставке
                
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


# --- Вспомогательные функции для работы с диаграммами ---

def _parse_title_ref(title_ref: str) -> Optional[Dict[str, Any]]:
    """
    Парсит строку ссылки на заголовок диаграммы (например, "=Sheet1!$A$1:$B$2")
    и возвращает словарь с информацией о листе и диапазоне.

    Args:
        title_ref (str): Строка ссылки на заголовок.

    Returns:
        Optional[Dict[str, Any]]: Словарь с ключами 'sheet_name', 'range_address',
                                  'is_single_cell' (bool), или None, если парсинг не удался.
    """
    if not title_ref or not title_ref.startswith("="):
        logger.warning(f"[ДИАГРАММА] Неверный формат title_ref: '{title_ref}'. Ожидается строка, начинающаяся с '='.")
        return None

    # Убираем знак '=' в начале
    ref_body = title_ref[1:]
    
    # Разделяем по '!'
    if '!' not in ref_body:
        logger.warning(f"[ДИАГРАММА] Неверный формат title_ref: '{title_ref}'. Ожидается 'SheetName!Range'.")
        return None

    sheet_part, range_part = ref_body.split('!', 1)
    sheet_name = sheet_part.strip("'") # Убираем одинарные кавычки, если они есть
    range_address = range_part

    # Определяем, является ли диапазон одной ячейкой
    is_single_cell = ':' not in range_address

    return {
        'sheet_name': sheet_name,
        'range_address': range_address,
        'is_single_cell': is_single_cell
    }

def _extract_single_value_from_ref(db_path: Union[str, Path], ref: str) -> Optional[str]:
    """
    Извлекает значение из ссылки на ячейку (например, "=Sheet1!$A$1").
    Использует ProjectDBStorage для загрузки данных.

    Args:
        db_path (Union[str, Path]): Путь к файлу БД проекта.
        ref (str): Строка ссылки на ячейку.

    Returns:
        Optional[str]: Значение ячейки или None в случае ошибки.
    """
    try:
        # Убираем начальный '='
        if ref.startswith('='):
            ref_body = ref[1:]
        else:
            ref_body = ref
        
        # Разделяем на имя листа и адрес ячейки
        if '!' not in ref_body:
            logger.error(f"[ДИАГРАММА] Неверный формат ссылки для извлечения значения: {ref}")
            return None
            
        sheet_name_part, cell_address = ref_body.split('!', 1)
        # Убираем одинарные кавычки из имени листа, если они есть
        sheet_name = sheet_name_part.strip("'")
        
        # Загружаем данные листа из БД
        storage = ProjectDBStorage(str(db_path))
        if not storage.connect():
            logger.error(f"[ДИАГРАММА] Не удалось подключиться к БД для извлечения значения из ссылки: {ref}")
            return None
        
        raw_data = storage.load_sheet_raw_data(sheet_name)
        storage.disconnect()
        
        # Ищем значение по адресу ячейки
        for item in raw_data:
            if item.get('cell_address') == cell_address:
                return str(item.get('value', ''))
                
        logger.warning(f"[ДИАГРАММА] Ячейка {cell_address} не найдена на листе {sheet_name} для ссылки: {ref}")
        return None
        
    except Exception as e:
        logger.error(f"[ДИАГРАММА] Ошибка при извлечении значения из ссылки {ref}: {e}", exc_info=True)
        return None

# Дополнительные вспомогательные функции можно добавить здесь
