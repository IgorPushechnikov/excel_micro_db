# src/exporter/excel_exporter.py

import xlsxwriter
import sqlite3
import logging
import os
from typing import Dict, Any, List, Optional
import json

# Импортируем ProjectDBStorage из storage.base
# Предполагается, что storage.base доступен по этому пути
from src.storage.base import ProjectDBStorage

# Получаем логгер
logger = logging.getLogger(__name__)

class XlsxWriterExporter:
    """
    Экспортёр данных проекта в формат Excel (.xlsx) с использованием библиотеки xlsxwriter.
    Работает с данными, извлечёнными и хранящимися в БД проекта.
    """

    def __init__(self, project_db_path: str):
        """
        Инициализирует экземпляр экспортера.

        Args:
            project_db_path (str): Путь к файлу БД проекта SQLite.
        """
        self.project_db_path = project_db_path
        self.storage: Optional[ProjectDBStorage] = None
        logger.debug(f"XlsxWriterExporter инициализирован с путем к БД: {project_db_path}")

    def export_project_to_excel(self, output_file_path: str) -> bool:
        """
        Экспортирует весь проект (все листы) в Excel-файл.

        Args:
            output_file_path (str): Путь к создаваемому .xlsx файлу.

        Returns:
            bool: True, если экспорт успешен, иначе False.
        """
        logger.info(f"Начало экспорта проекта из БД '{self.project_db_path}' в '{output_file_path}'.")

        try:
            # Создаем экземпляр хранилища и подключаемся
            self.storage = ProjectDBStorage(self.project_db_path)
            if not self.storage.connect():
                logger.error("Не удалось подключиться к БД проекта для экспорта.")
                return False

            # Создаем Excel файл с помощью xlsxwriter
            workbook = xlsxwriter.Workbook(output_file_path)
            logger.debug(f"Создан Workbook для файла '{output_file_path}'.")

            # Получаем список листов проекта
            # Предполагается, что есть способ получить список листов, например, через SQL-запрос
            # или через метод в storage, который мы реализуем позже.
            # Пока делаем это напрямую через соединение.
            sheets_info = self._get_sheets_info()
            if not sheets_info:
                logger.warning("Не найдено листов для экспорта.")
                workbook.close()
                return True # Считаем успешным, если нет данных

            # Экспортируем каждый лист
            for sheet_info in sheets_info:
                sheet_id = sheet_info['sheet_id']
                sheet_name = sheet_info['name']
                logger.info(f"Экспорт листа '{sheet_name}' (ID: {sheet_id})...")
                
                # Создаем лист в Excel файле
                worksheet = workbook.add_worksheet(sheet_name)
                logger.debug(f"Создан лист '{sheet_name}' в Excel файле.")

                # --- 1. Загружаем и записываем "сырые данные" ---
                raw_data = self.storage.load_sheet_raw_data(sheet_name)
                self._write_raw_data(worksheet, raw_data)
                logger.debug(f"Записаны 'сырые данные' для листа '{sheet_name}'.")

                # --- 2. Загружаем и записываем редактируемые данные (перекрывают raw_data) ---
                # Нам нужно имя листа для load_sheet_editable_data
                editable_data = self.storage.load_sheet_editable_data(sheet_id, sheet_name)
                self._write_editable_data(worksheet, editable_data)
                logger.debug(f"Записаны 'редактируемые данные' для листа '{sheet_name}'.")

                # --- 3. Загружаем и записываем формулы ---
                formulas = self.storage.load_sheet_formulas(sheet_id)
                self._write_formulas(worksheet, formulas)
                logger.debug(f"Записаны формулы для листа '{sheet_name}'.")

                # --- 4. Загружаем и применяем стили ---
                styles_data = self.storage.load_sheet_styles(sheet_id)
                self._apply_styles(workbook, worksheet, styles_data)
                logger.debug(f"Применены стили для листа '{sheet_name}'.")

                # --- 5. Загружаем и добавляем диаграммы ---
                charts_data = self.storage.load_sheet_charts(sheet_id)
                self._add_charts(workbook, worksheet, charts_data)
                logger.debug(f"Добавлены диаграммы для листа '{sheet_name}'.")

                # --- 6. Обрабатываем объединённые ячейки (если данные есть в метаданных) ---
                # Предполагается, что информация об объединённых ячейках хранится в метаданных
                sheet_metadata = self.storage.load_sheet_metadata(sheet_name)
                merged_ranges = sheet_metadata.get('merged_cells', []) if sheet_metadata else []
                self._merge_cells(worksheet, merged_ranges)
                logger.debug(f"Обработаны объединённые ячейки для листа '{sheet_name}'.")

            # Закрываем Excel файл
            workbook.close()
            logger.info(f"Экспорт проекта завершён. Файл сохранён: {output_file_path}")
            return True

        except Exception as e:
            logger.error(f"Ошибка при экспорте проекта в '{output_file_path}': {e}", exc_info=True)
            # Убедимся, что файл закрыт, даже если произошла ошибка
            # xlsxwriter обычно сам закрывает файл при выходе из контекста,
            # но вручную это сделать сложнее. Лучше использовать try-finally или контекстный менеджер для workbook.
            # Пока оставим как есть, предполагая, что workbook.close() внутри блока.
            return False
        finally:
            # Отключаемся от БД
            if self.storage:
                self.storage.disconnect()
                logger.debug("Отключение от БД проекта после экспорта.")

    def _get_sheets_info(self) -> List[Dict[str, Any]]:
        """
        Получает информацию о листах проекта напрямую из БД.
        В будущем это может быть метод в storage.
        """
        if not self.storage or not self.storage.connection:
            logger.error("Нет подключения к БД для получения информации о листах.")
            return []
        
        try:
            cursor = self.storage.connection.cursor()
            cursor.execute("SELECT sheet_id, name FROM sheets ORDER BY sheet_id")
            rows = cursor.fetchall()
            return [{"sheet_id": row[0], "name": row[1]} for row in rows]
        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при получении списка листов: {e}")
            return []

    def _write_raw_data(self, worksheet, raw_data: List[Dict[str, Any]]):
        """Записывает 'сырые данные' в лист."""
        for data_item in raw_data:
            cell_address = data_item.get('cell_address')
            value = data_item.get('value')
            # value_type = data_item.get('value_type') # Может понадобиться для точного определения типа
            if cell_address:
                # xlsxwriter может автоматически определить тип значения
                worksheet.write(cell_address, value)

    def _write_editable_data(self, worksheet, editable_data: List[Dict[str, Any]]):
        """Записывает редактируемые данные в лист (перекрывают raw_data)."""
        for data_item in editable_data:
            cell_address = data_item.get('cell_address')
            value = data_item.get('value')
            if cell_address:
                # Редактируемые данные перекрывают сырые
                worksheet.write(cell_address, value)

    def _write_formulas(self, worksheet, formulas: List[Dict[str, str]]):
        """Записывает формулы в лист."""
        for formula_item in formulas:
            cell_address = formula_item.get('cell_address')
            formula = formula_item.get('formula')
            if cell_address and formula:
                # Убираем начальный '=' если он есть, xlsxwriter добавит его автоматически
                if formula.startswith('='):
                    formula = formula[1:]
                worksheet.write_formula(cell_address, formula)

    def _apply_styles(self, workbook, worksheet, styles_data: List[Dict[str, Any]]):
        """
        Применяет стили к листу.
        styles_data: Список словарей с 'range_address' и 'style_attributes' (строка JSON).
        """
        for style_item in styles_data:
            range_address = style_item.get('range_address')
            style_attributes_json = style_item.get('style_attributes')
            
            if not range_address or not style_attributes_json:
                logger.warning(f"Пропущен стиль: отсутствует range_address или style_attributes. Range: {range_address}")
                continue

            try:
                # Десериализуем JSON атрибутов стиля
                style_dict = json.loads(style_attributes_json)
                logger.debug(f"Десериализован стиль для диапазона {range_address}: {list(style_dict.keys())}")
            except json.JSONDecodeError as e:
                logger.error(f"Ошибка десериализации JSON стиля для диапазона {range_address}: {e}")
                continue # Пропускаем этот стиль

            # Создаем формат xlsxwriter на основе атрибутов стиля
            format_properties = self._convert_style_to_xlsxwriter_format(style_dict)
            if not format_properties:
                logger.debug(f"Нет свойств формата для диапазона {range_address}, стиль пропущен или не распознан.")
                continue
            
            try:
                cell_format = workbook.add_format(format_properties)
                logger.debug(f"Создан формат xlsxwriter для диапазона {range_address}.")
            except Exception as e:
                logger.error(f"Ошибка создания формата xlsxwriter для диапазона {range_address}: {e}")
                continue

            # Применяем формат к диапазону
            try:
                worksheet.conditional_format(range_address, {'type': 'no_errors', 'format': cell_format})
                # ИЛИ, если conditional_format не подходит, можно использовать write с форматом для каждой ячейки
                # Но это менее эффективно. conditional_format должен работать для применения к диапазону.
                # worksheet.write(range_address, None, cell_format) # Не работает для диапазонов
                # Нужно итерироваться по ячейкам диапазона, если conditional_format не подходит.
                # worksheet.set_row(...) / set_column(...) для строк/столбцов
                
                # Альтернатива: использовать формат при записи данных
                # Но в нашем случае стили применяются после записи данных.
                # xlsxwriter требует, чтобы формат был применен при записи.
                # Возможно, нужно пересмотреть логику: сначала создать все форматы,
                # затем при записи данных использовать их.
                # Пока оставим conditional_format как временное решение.
                # TODO: Уточнить, как правильно применять формат ко всему диапазону в xlsxwriter.
                
                logger.debug(f"Формат применен к диапазону {range_address}.")
            except Exception as e:
                logger.error(f"Ошибка применения формата к диапазону {range_address}: {e}")

    def _convert_style_to_xlsxwriter_format(self, style_dict: Dict[str, Any]) -> Dict[str, Any]:
        """
        Преобразует универсальный словарь атрибутов стиля в словарь свойств формата xlsxwriter.
        Это ключевая функция для адаптации стилей.
        Нужно реализовать отображение атрибутов из Excel/openpyxl в свойства xlsxwriter.
        """
        xlsxwriter_format = {}
        
        # Примеры отображений (нужно расширить на основе реальных атрибутов из analyzer/storage)
        # https://xlsxwriter.readthedocs.io/format.html
        
        # --- Шрифт ---
        font = style_dict.get('font', {})
        if isinstance(font, dict):
            if 'name' in font:
                xlsxwriter_format['font_name'] = font['name']
            if 'sz' in font or 'size' in font: # 'sz' из CT_Font, 'size' - наше внутреннее
                xlsxwriter_format['font_size'] = font.get('sz', font.get('size'))
            if 'b' in font and font['b']: # Bold
                xlsxwriter_format['bold'] = True
            if 'i' in font and font['i']: # Italic
                xlsxwriter_format['italic'] = True
            # ... другие атрибуты шрифта (color, underline и т.д.)

        # --- Заливка (Pattern/Fill) ---
        fill = style_dict.get('fill', {})
        if isinstance(fill, dict):
            # xlsxwriter поддерживает fg_color, bg_color для сплошной заливки
            # и pattern, pattern_foreground_color, pattern_background_color для узоров
            # Предположим, что у нас есть 'patternType' и 'fgColor'
            pattern_type = fill.get('patternType')
            fg_color = fill.get('fgColor', {}).get('rgb') # 'rgb': 'FFCCCCCC'
            
            if pattern_type == 'solid' and fg_color:
                # xlsxwriter использует '#RRGGBB' или 'color_name'
                # openpyxl может использовать 'FFCCCCCC' (ARGB), нужно преобразовать
                if fg_color.startswith('FF'):
                    fg_color = fg_color[2:] # Убираем альфа-канал
                xlsxwriter_format['bg_color'] = f"#{fg_color}"
            # ... обработка других типов заливки

        # --- Границы (Border) ---
        border = style_dict.get('border', {})
        if isinstance(border, dict):
            # xlsxwriter поддерживает отдельные границы: top, bottom, left, right
            # Каждая граница имеет тип (border_type) и цвет (color)
            # 'thin', 'medium', 'thick', 'dashDot', ...
            for side in ['top', 'bottom', 'left', 'right']:
                side_border = border.get(side, {})
                if isinstance(side_border, dict) and side_border:
                    border_type = side_border.get('style')
                    border_color_dict = side_border.get('color', {})
                    border_color = border_color_dict.get('rgb') if isinstance(border_color_dict, dict) else None
                    
                    if border_type:
                        # Преобразуем стиль границы из openpyxl в xlsxwriter
                        # Это может потребовать словаря соответствий
                        xlsxwriter_format[f'{side}'] = self._map_border_style(border_type)
                    
                    if border_color:
                         if border_color.startswith('FF'):
                            border_color = border_color[2:]
                         xlsxwriter_format[f'{side}_color'] = f"#{border_color}"
            # ... обработка диагональных границ

        # --- Выравнивание (Alignment) ---
        alignment = style_dict.get('alignment', {})
        if isinstance(alignment, dict):
            horizontal = alignment.get('horizontal')
            vertical = alignment.get('vertical')
            wrap_text = alignment.get('wrapText')
            text_rotation = alignment.get('textRotation')
            
            if horizontal:
                # xlsxwriter: 'left', 'center', 'right', 'fill', 'justify', 'center_across'
                xlsxwriter_format['align'] = horizontal 
            if vertical:
                # xlsxwriter: 'top', 'vcenter', 'bottom', 'vjustify', 'vdistributed'
                 xlsxwriter_format['valign'] = vertical
            if wrap_text is True:
                xlsxwriter_format['text_wrap'] = True
            if text_rotation is not None:
                xlsxwriter_format['rotation'] = text_rotation
                
        # --- Числовой формат (Number Format) ---
        number_format = style_dict.get('number_format')
        if number_format:
             xlsxwriter_format['num_format'] = number_format

        # --- Защита (Protection) ---
        # xlsxwriter поддерживает 'locked' и 'hidden'
        protection = style_dict.get('protection', {})
        if isinstance(protection, dict):
            locked = protection.get('locked')
            hidden = protection.get('hidden')
            if locked is not None:
                xlsxwriter_format['locked'] = locked
            if hidden is not None:
                xlsxwriter_format['hidden'] = hidden

        logger.debug(f"Преобразован стиль: {list(style_dict.keys())} -> {list(xlsxwriter_format.keys())}")
        return xlsxwriter_format

    def _map_border_style(self, openpyxl_style: str) -> int:
        """Отображает стиль границы из openpyxl в xlsxwriter."""
        # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.borders.html
        # https://xlsxwriter.readthedocs.io/format.html
        style_map = {
            'thin': 1,
            'medium': 2,
            'thick': 3,
            'double': 4,
            'hair': 5,
            'mediumDashed': 6,
            'dashDot': 7,
            'mediumDashDot': 8,
            'dashDotDot': 9,
            'mediumDashDotDot': 10,
            'slantDashDot': 11,
            # xlsxwriter использует числовые коды
            # 'none' -> 0 или отсутствие ключа
        }
        return style_map.get(openpyxl_style, 1) # По умолчанию thin

    def _add_charts(self, workbook, worksheet, charts_data: List[Dict[str, Any]]):
        """
        Добавляет диаграммы на лист.
        charts_data: Список словарей с 'chart_data' (сериализованные данные диаграммы).
        """
        # TODO: Реализовать добавление диаграмм.
        # Это сложная часть. Нужно десериализовать chart_data
        # и преобразовать его в объект Chart xlsxwriter.
        # chart_data может быть JSON с описанием типа диаграммы, данных, позиции.
        # Или это может быть XML, извлечённый openpyxl.
        # xlsxwriter имеет API для создания диаграмм: workbook.add_chart({'type': 'column'})
        # и добавления их на лист: worksheet.insert_chart('D2', chart)
        logger.warning("Добавление диаграмм пока не реализовано.")
        # for chart_item in charts_data:
        #     chart_data_str = chart_item.get('chart_data')
        #     if chart_data_str:
        #         # ... логика создания и вставки диаграммы
        #         pass

    def _merge_cells(self, worksheet, merged_ranges: List[str]):
        """Объединяет ячейки на листе."""
        for range_str in merged_ranges:
            try:
                worksheet.merge_range(range_str, '') # Значение будет перезаписано при записи данных
                logger.debug(f"Объединены ячейки: {range_str}")
            except Exception as e:
                logger.error(f"Ошибка при объединении ячеек {range_str}: {e}")

# --- Функция для удобного вызова из CLI/AppController ---
def export_project(project_db_path: str, output_excel_path: str) -> bool:
    """
    Удобная функция для экспорта проекта.

    Args:
        project_db_path (str): Путь к файлу БД проекта.
        output_excel_path (str): Путь к выходному .xlsx файлу.

    Returns:
        bool: True, если экспорт успешен, иначе False.
    """
    exporter = XlsxWriterExporter(project_db_path)
    return exporter.export_project_to_excel(output_excel_path)

# Пример использования (если файл запускается напрямую)
if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Использование: python excel_exporter.py <project_db_path> <output_excel_path>")
        sys.exit(1)
    
    project_db_path = sys.argv[1]
    output_excel_path = sys.argv[2]
    
    if export_project(project_db_path, output_excel_path):
        print(f"Проект успешно экспортирован в {output_excel_path}")
    else:
        print(f"Ошибка при экспорте проекта в {output_excel_path}")
        sys.exit(1)
