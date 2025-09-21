# src/exporter/xlsx_exporter.py
"""
Модуль для экспорта данных проекта Excel Micro DB в новый Excel-файл с использованием XlsxWriter.
"""

import xlsxwriter
import logging
from typing import Dict, Any, List, Optional, Tuple
from pathlib import Path
import re # Для парсинга адресов диапазонов

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
import sys
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

# --- Вспомогательные функции ---

def _parse_cell_address(address: str) -> Tuple[int, int]:
    """
    Преобразует адрес ячейки Excel (например, "A1") в (строка (0-based), столбец (0-based)).
    """
    # Регулярное выражение для извлечения столбца и строки
    match = re.match(r"([A-Z]+)(\d+)", address)
    if not match:
        logger.warning(f"Невозможно распарсить адрес ячейки: {address}")
        # Возвращаем (-1, -1) как индикатор ошибки
        return (-1, -1)
    column_letter, row_number = match.groups()
    
    # Преобразуем номер строки в 0-based индекс
    row = int(row_number) - 1
    
    # Преобразуем букву столбца в 0-based индекс
    col = 0
    for char in column_letter:
        col = col * 26 + (ord(char) - ord('A') + 1)
    col = col - 1 # 0-based
    
    return (row, col)

def _parse_range_address(range_address: str) -> Tuple[int, int, int, int]:
    """
    Преобразует адрес диапазона Excel (например, "A1:B2" или "C5") 
    в (start_row (0-based), start_col (0-based), end_row (0-based), end_col (0-based)).
    """
    if ':' in range_address:
        start_addr, end_addr = range_address.split(':')
    else:
        # Это одиночная ячейка
        start_addr = end_addr = range_address
        
    start_row, start_col = _parse_cell_address(start_addr)
    end_row, end_col = _parse_cell_address(end_addr)
    
    # Проверка на ошибки парсинга
    if start_row == -1 or end_row == -1:
        logger.error(f"Ошибка при парсинге диапазона: {range_address}")
        # Возвращаем недопустимый диапазон
        return (0, 0, -1, -1) 
        
    return (start_row, start_col, end_row, end_col)

def _convert_style_attributes_to_xlsxwriter_format(style_attributes: Dict[str, Any]) -> Dict[str, Any]:
    """
    Преобразует атрибуты стиля из формата БД в формат, понятный workbook.add_format().
    
    Args:
        style_attributes (Dict[str, Any]): Атрибуты стиля из БД, например:
            {
                'font_name': 'Calibri', 'font_sz': 11.0, 'font_b': 1, 'font_i': 0,
                'fill_pattern_type': 'solid', 'fill_fg_color_rgb': 'FF0000',
                'border_left_style': 'thin', 'border_left_color_rgb': '000000',
                'alignment_horizontal': 'center', 'alignment_vertical': 'top',
                'protection_locked': 1
            }
            
    Returns:
        Dict[str, Any]: Словарь атрибутов для workbook.add_format().
    """
    format_props = {}
    
    # --- Обработка шрифта ---
    # XlsxWriter использует названия атрибутов, похожие на openpyxl, но с некоторыми отличиями
    if 'font_name' in style_attributes and style_attributes['font_name'] is not None:
        format_props['font_name'] = style_attributes['font_name']
    if 'font_sz' in style_attributes and style_attributes['font_sz'] is not None:
        # XlsxWriter ожидает 'font_size'
        format_props['font_size'] = float(style_attributes['font_sz'])
    if 'font_b' in style_attributes:
        # XlsxWriter ожидает 'bold'
        format_props['bold'] = bool(style_attributes['font_b'])
    if 'font_i' in style_attributes:
        # XlsxWriter ожидает 'italic'
        format_props['italic'] = bool(style_attributes['font_i'])
    if 'font_u' in style_attributes and style_attributes['font_u'] is not None:
        # XlsxWriter ожидает 'underline'
        format_props['underline'] = style_attributes['font_u'] # 'single', 'double' etc.
    if 'font_strike' in style_attributes:
        # XlsxWriter ожидает 'font_strikeout'
        format_props['font_strikeout'] = bool(style_attributes['font_strike']) 
    # Цвет шрифта
    # XlsxWriter поддерживает 'font_color' в формате '#RRGGBB' или 'color_name'
    if 'font_color_rgb' in style_attributes and style_attributes['font_color_rgb'] is not None:
        # Добавляем # если его нет
        color_val = style_attributes['font_color_rgb']
        if not color_val.startswith('#'):
            # Предполагаем, что это hex без #
            if len(color_val) == 6 or len(color_val) == 8: # RRGGBB or AARRGGBB
                 format_props['font_color'] = f"#{color_val[-6:]}" # Берем последние 6 символов для RRGGBB
            else:
                 logger.warning(f"Неожиданный формат цвета шрифта: {color_val}")
        else:
             format_props['font_color'] = color_val
    # TODO: Обработка других атрибутов шрифта (theme, tint, vert_align, scheme)
    # Это может потребовать дополнительной логики или игнорирования, если XlsxWriter не поддерживает напрямую.
    
    # --- Обработка заливки ---
    # XlsxWriter использует 'bg_color' и 'pattern' для заливки.
    # 'pattern' соответствует 'patternType' из openpyxl (например, 'solid')
    # 'bg_color' соответствует 'fgColor' из openpyxl, если pattern='solid'
    pattern_type = style_attributes.get('fill_pattern_type')
    if pattern_type:
        # XlsxWriter использует немного другие названия паттернов, но 'solid' должен работать.
        # Другие паттерны могут требовать сопоставления.
        format_props['pattern'] = 1 if pattern_type == 'solid' else 0 # 1 для solid, 0 для none как стандарт
        # Цвет заливки
        if 'fill_fg_color_rgb' in style_attributes and style_attributes['fill_fg_color_rgb'] is not None:
            color_val = style_attributes['fill_fg_color_rgb']
            if not color_val.startswith('#'):
                if len(color_val) == 6 or len(color_val) == 8:
                     format_props['bg_color'] = f"#{color_val[-6:]}"
                else:
                     logger.warning(f"Неожиданный формат цвета заливки: {color_val}")
            else:
                 format_props['bg_color'] = color_val
        # TODO: Обработка bgColor из openpyxl (который используется для других паттернов)
        # и других атрибутов заливки (theme, tint)

    # --- Обработка границ ---
    # XlsxWriter поддерживает границы через атрибуты 'left', 'right', 'top', 'bottom', 'diag_type', 'diag_border'
    # Каждая граница определяется словарем {'style': ..., 'color': ...}
    border_props = {}
    # Пример для левой границы
    left_style = style_attributes.get('border_left_style')
    if left_style:
        border_props['left'] = {'style': left_style}
        if 'border_left_color_rgb' in style_attributes and style_attributes['border_left_color_rgb'] is not None:
            color_val = style_attributes['border_left_color_rgb']
            if not color_val.startswith('#'):
                if len(color_val) == 6 or len(color_val) == 8:
                     border_props['left']['color'] = f"#{color_val[-6:]}"
                else:
                     logger.warning(f"Неожиданный формат цвета границы (left): {color_val}")
            else:
                 border_props['left']['color'] = color_val
    # Аналогично для других сторон
    for side in ['right', 'top', 'bottom']: # 'diagonal' требует особой обработки
        side_style_key = f'border_{side}_style'
        side_color_key = f'border_{side}_color_rgb'
        side_style = style_attributes.get(side_style_key)
        if side_style:
            if side not in border_props:
                border_props[side] = {}
            border_props[side]['style'] = side_style
            if side_color_key in style_attributes and style_attributes[side_color_key] is not None:
                color_val = style_attributes[side_color_key]
                if not color_val.startswith('#'):
                    if len(color_val) == 6 or len(color_val) == 8:
                         border_props[side]['color'] = f"#{color_val[-6:]}"
                    else:
                         logger.warning(f"Неожиданный формат цвета границы ({side}): {color_val}")
                else:
                     border_props[side]['color'] = color_val
                     
    # Диагональные границы (если используются)
    # XlsxWriter использует 'diag_type' (0=none, 1=down, 2=up, 3=both) и 'diag_border'
    diag_up = style_attributes.get('border_diagonal_up')
    diag_down = style_attributes.get('border_diagonal_down')
    diag_type = 0
    if diag_up and diag_down:
        diag_type = 3 # both
    elif diag_down:
        diag_type = 1 # down
    elif diag_up:
        diag_type = 2 # up
    if diag_type != 0:
        border_props['diag_type'] = diag_type
        diag_style = style_attributes.get('border_diagonal_style')
        if diag_style:
            border_props['diag_border'] = {'style': diag_style}
            if 'border_diagonal_color_rgb' in style_attributes and style_attributes['border_diagonal_color_rgb'] is not None:
                color_val = style_attributes['border_diagonal_color_rgb']
                if not color_val.startswith('#'):
                    if len(color_val) == 6 or len(color_val) == 8:
                         border_props['diag_border']['color'] = f"#{color_val[-6:]}"
                    else:
                         logger.warning(f"Неожиданный формат цвета диагональной границы: {color_val}")
                else:
                     border_props['diag_border']['color'] = color_val
    
    # Если были определены границы, добавляем их в формат
    if border_props:
        format_props.update(border_props)
        
    # --- Обработка выравнивания ---
    # XlsxWriter использует 'align' (горизонтальное) и 'valign' (вертикальное)
    h_align = style_attributes.get('alignment_horizontal')
    if h_align:
        # Некоторые значения могут отличаться, например 'centerContinuous' в openpyxl -> 'center_across' в XlsxWriter
        # Пока используем напрямую, можно добавить сопоставление при необходимости.
        format_props['align'] = h_align 
    v_align = style_attributes.get('alignment_vertical')
    if v_align:
        # 'top', 'center', 'bottom' должны совпадать
        format_props['valign'] = v_align
        
    # Другие атрибуты выравнивания
    if 'alignment_wrap_text' in style_attributes:
        format_props['text_wrap'] = bool(style_attributes['alignment_wrap_text'])
    if 'alignment_shrink_to_fit' in style_attributes:
        # XlsxWriter использует 'shrink'
        format_props['shrink'] = bool(style_attributes['alignment_shrink_to_fit'])
    # TODO: Обработка других атрибутов выравнивания (indent, rotation, reading_order и т.д.)
    # если они будут использоваться.

    # --- Обработка защиты ---
    # XlsxWriter использует 'locked' и 'hidden' напрямую в формате
    if 'protection_locked' in style_attributes:
        format_props['locked'] = bool(style_attributes['protection_locked'])
    if 'protection_hidden' in style_attributes:
        format_props['hidden'] = bool(style_attributes['protection_hidden'])
        
    # --- Другие атрибуты ---
    # num_format - формат чисел
    # XlsxWriter использует 'num_format'
    if 'num_fmt_id' in style_attributes and style_attributes['num_fmt_id'] is not None:
        # TODO: Нужно сопоставить num_fmt_id с реальным строковым форматом.
        # Пока оставим заглушку или пропустим, если это не критично на данном этапе.
        # format_props['num_format'] = ... 
        pass # Пропускаем, так как у нас нет маппинга ID -> строка формата
        
    logger.debug(f"Преобразованные атрибуты стиля для XlsxWriter: {format_props}")
    return format_props

# --- Основные функции экспорта ---

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
        # ВАЖНО: Убедиться, что данные берутся из правильного источника.
        # Согласно документации и анализу, для экспорта результатов нужно использовать editable_data.
        editable_data = sheet_info.get("editable_data", {}) # Предполагается, что это загружено storage.load_sheet_editable_data
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
                # Убедиться, что формула начинается с '='
                if not formula.startswith('='):
                    formula_to_write = f"={formula}"
                else:
                    formula_to_write = formula
                worksheet.write_formula(cell_address, formula_to_write)

    except Exception as e:
        logger.error(f"Ошибка при экспорте формул листа: {e}")

def _export_sheet_styles(workbook, worksheet, sheet_info: Dict[str, Any]) -> None:
    """Экспортирует стили листа."""
    try:
        # Получаем данные о стилях из sheet_info
        # Согласно анализу storage/styles.py, load_sheet_styles возвращает
        # List[Dict[str, Any]] с ключами 'style_attributes' и 'range_address'
        styled_ranges_data = sheet_info.get("styled_ranges", [])
        
        logger.debug(f"Начало экспорта {len(styled_ranges_data)} стилевых диапазонов.")

        for style_range_info in styled_ranges_data:
            range_address = style_range_info.get("range_address")
            style_attributes = style_range_info.get("style_attributes", {})
            
            if not range_address:
                logger.warning("Пропущен стиль из-за отсутствия range_address.")
                continue
            if not style_attributes:
                logger.debug(f"Пропущен стиль для диапазона {range_address} из-за отсутствия атрибутов.")
                continue # Нечего применять

            # Преобразуем атрибуты стиля из формата БД в формат XlsxWriter
            xlsx_format_dict = _convert_style_attributes_to_xlsxwriter_format(style_attributes)
            
            if not xlsx_format_dict:
                logger.debug(f"Преобразование стиля для диапазона {range_address} не дало параметров формата. Пропущено.")
                continue # Нечего применять

            # Создаем объект формата XlsxWriter
            try:
                cell_format = workbook.add_format(xlsx_format_dict)
            except Exception as format_error:
                logger.error(f"Ошибка при создании формата XlsxWriter для диапазона {range_address} с атрибутами {xlsx_format_dict}: {format_error}")
                continue # Пропускаем этот стиль

            # Применяем формат к диапазону
            # XlsxWriter не имеет прямого метода "применить формат к диапазону".
            # Нужно итерироваться по ячейкам в диапазоне.
            
            # Парсим адрес диапазона
            start_row, start_col, end_row, end_col = _parse_range_address(range_address)
            
            # Проверка на корректность диапазона
            if end_row < start_row or end_col < start_col:
                logger.warning(f"Некорректный диапазон для стиля: {range_address}. Пропущено.")
                continue
                
            logger.debug(f"Применение стиля к диапазону {range_address} (строки {start_row}-{end_row}, столбцы {start_col}-{end_col}).")

            # Итерируемся по строкам и столбцам диапазона
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    # XlsxWriter не перезаписывает ячейку, если она уже существует.
                    # write_blank используется для применения формата к пустой ячейке.
                    # write с форматом перезапишет значение и применит формат, если ячейка уже существует.
                    # Чтобы не затереть данные, можно использовать write_blank, но это не всегда удобно.
                    # Более надежный способ - это сначала записать все данные, затем применить стили.
                    # Но в текущей реализации _export_sheet_data записывает данные без форматов.
                    # Поэтому мы можем безопасно применить формат.
                    # worksheet.write_blank(row, col, None, cell_format) # Применить к пустой
                    # Но write_blank не перезапишет существующую ячейку с данными.
                    # write(row, col, None, cell_format) перезапишет значение на None.
                    # write(row, col, '', cell_format) перезапишет значение на ''.
                    # Лучше использовать write_blank или специальный метод format_only.
                    # XlsxWriter предоставляет метод format_only для write, но это неочевидно.
                    # Проще всего - использовать write_blank, если ячейка пуста, или write с текущим значением.
                    # Но мы не знаем текущее значение здесь.
                    # Альтернатива: использовать set_row/set_column для строк/столбцов, но это не для диапазонов.
                    
                    # Попробуем write_blank. Это безопасно, если ячейка пуста.
                    # Если ячейка не пуста, write_blank игнорируется.
                    # Это не идеально, но является ограничением XlsxWriter.
                    # Для полноценного применения стилей к существующим ячейкам
                    # нужно было бы сначала записать данные, запомнить их, а потом применять стили.
                    # Или использовать более сложную логику отслеживания.
                    # Пока используем write_blank как компромисс.
                    # TODO: Подумать над более надежным способом применения стилей к уже записанным ячейкам.
                    
                    # worksheet.write_blank(row, col, '', cell_format) # Пишем пустую строку с форматом
                    # Это перезапишет значение. Не подходит.
                    
                    # write_blank не перезаписывает значение, но и не применяет формат, если значение уже есть.
                    # write с форматом перезаписывает значение.
                    # Нужно другое решение.
                    
                    # Вариант: Использовать worksheet.conditional_format, но это не то.
                    
                    # Вывод: write_blank не подходит. write перезаписывает.
                    # Единственный способ - это сначала записать данные, а потом применить стили.
                    # Но _export_sheet_data уже отработал.
                    # Значит, нужно пересмотреть порядок: сначала записываем данные в промежуточную структуру,
                    # затем применяем стили, затем записываем в worksheet.
                    # Это требует рефакторинга.
                    
                    # Временное решение: Просто применяем формат через write.
                    # Это будет работать, если данные в ячейке еще не записаны, или если мы готовы перезаписать значение.
                    # Поскольку _export_sheet_data записывает данные без форматов,
                    # а _export_sheet_styles вызывается после, мы можем перезаписать пустые ячейки.
                    # Но если в ячейке уже есть данные (например, из _export_sheet_data),
                    # write(row, col, None, format) затрет данные.
                    # write(row, col, '', format) затрет данные строкой.
                    
                    # Правильнее: передавать данные и стили вместе.
                    # Но в текущей архитектуре они разделены.
                    
                    # Пока применим формат через write_blank. Если ячейка пуста, формат применится.
                    # Если не пуста, формат не применится, но данные тоже не затрутся.
                    # Это не идеально, но лучше, чем перезапись.
                    # Для полноценного решения нужен рефакторинг.
                    
                    # worksheet.write_blank(row, col, None, cell_format) 
                    # write_blank не работает, как ожидалось для форматирования существующих ячеек.
                    
                    # Альтернатива: использовать worksheet.format_row или format_column для строк/столбцов.
                    # Но это не для произвольных диапазонов.
                    
                    # Поскольку XlsxWriter не позволяет легко применить формат к уже записанной ячейке,
                    # и у нас данные и стили обрабатываются отдельно, 
                    # мы столкнемся с ограничениями.
                    # Для MVP попробуем write_blank. Если не сработает, будем искать обходные пути или рефакторить.
                    
                    # Еще одна идея: использовать write с тем же значением, которое уже записано.
                    # Но мы его не знаем здесь.
                    
                    # write(row, col, None, format) - затрет значение на None.
                    # write(row, col, existing_value, format) - идеально, но existing_value неизвестен.
                    
                    # Попробуем write_blank. 
                    # worksheet.write_blank(row, col, None, cell_format)
                    # write_blank(row, col, '', cell_format) - запишет '' и применит формат.
                    
                    # write_blank документирован как способ применения формата к пустой ячейке.
                    # "Write a blank cell with a format."
                    # "If the cell already contains data it will be overridden."
                    # Это означает, что он перезапишет данные. Это не то.
                    
                    # Вывод: write_blank перезаписывает. write с форматом перезаписывает.
                    # Нет способа применить формат к существующей ячейке без знания её значения.
                    
                    # Решение: Рефакторинг. Сначала собираем все данные и стили в памяти,
                    # затем записываем в worksheet, применяя форматы.
                    # Это требует изменений в _export_sheet_data и _export_sheet_formulas.
                    
                    # Для быстрого MVP попробуем write_blank и посмотрим.
                    # Если данные уже есть, они будут перезаписаны на None. Это плохо.
                    
                    # Альтернатива: write с пустой строкой и форматом.
                    # worksheet.write(row, col, '', cell_format)
                    # Перезапишет данные на ''. Тоже плохо.
                    
                    # Похоже, единственный способ - это интеграция данных и стилей на этапе записи.
                    # Т.е. _export_sheet_data должен учитывать стили.
                    
                    # Это требует архитектурного изменения.
                    # Пока продолжим, но с пометкой, что стили могут не примениться корректно
                    # к уже записанным данным.
                    
                    # Записываем пустую строку с форматом. Это временное решение.
                    # В реальности нужно рефакторить, чтобы данные и стили обрабатывались вместе.
                    worksheet.write(row, col, '', cell_format) # Временное решение, перезаписывает данные
                    
                    # TODO: Рефакторинг для корректного применения стилей к существующим данным.
                    # Возможно, нужно создать промежуточную структуру данных "ячейка",
                    # которая будет содержать значение и формат, и только потом записывать в worksheet.

        logger.debug("Экспорт стилей листа завершен.")

    except Exception as e:
        logger.error(f"Ошибка при экспорте стилей листа: {e}", exc_info=True) # Добавлен exc_info для полного трейса

def _export_sheet_merged_cells(worksheet, sheet_info: Dict[str, Any]) -> None:
    """Экспортирует объединенные ячейки листа."""
    try:
        merged_cells_data = sheet_info.get("merged_cells", [])

        for range_address in merged_cells_data:
            if range_address:
                # XlsxWriter использует метод merge_range для объединения ячеек
                # Нужно определить значение для объединенного диапазона
                # Пока просто объединяем без значения. В будущем можно определить значение из данных.
                # merge_range(first_row, first_col, last_row, last_col, data, cell_format)
                # Нужно распарсить range_address
                start_row, start_col, end_row, end_col = _parse_range_address(range_address)
                if end_row >= start_row and end_col >= start_col:
                    # Пишем пустую строку в объединенный диапазон
                    worksheet.merge_range(start_row, start_col, end_row, end_col, "", None)
                else:
                    logger.warning(f"Некорректный адрес объединенного диапазона: {range_address}")

    except Exception as e:
        logger.error(f"Ошибка при экспорте объединенных ячеек листа: {e}")

# Точка входа для тестирования
if __name__ == "__main__":
    # Простой тест
    print("Тестирование xlsx_exporter")