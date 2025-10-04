# backend/constructor/widgets/new_gui/qt_model_adapter.py
"""
Модуль-адаптер для связи между AppController/БД и QAbstractTableModel (PySide6).
Предоставляет данные из БД для отображения в QTableView.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union # <-- Добавлен Union
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex, QPersistentModelIndex, QSize # <-- Добавлены QPersistentModelIndex, QSize
from PySide6.QtGui import QFont, QColor, QBrush, QTextOption

# Импортируем AppController
from backend.core.app_controller import create_app_controller
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class DBTableModel(QAbstractTableModel):
    """
    Модель данных для QTableView, получающая данные из AppController/БД.
    """

    def __init__(self, app_controller, sheet_name: str, parent=None):
        """
        Инициализирует модель.

        Args:
            app_controller: Экземпляр AppController.
            sheet_name (str): Имя листа для отображения.
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self.app_controller = app_controller
        self.sheet_name = sheet_name
        self._data: List[List[Any]] = []  # Список списков: строки x колонки
        self._headers: List[str] = []     # Заголовки колонок (A, B, C...)
        self._row_headers: List[str] = [] # Заголовки строк (1, 2, 3...)
        self._styles: Dict[tuple, Dict[str, Any]] = {} # Стили ячеек: {(row, col): {'font': ..., 'bg_color': ...}}
        self._merged_cells: List[tuple] = [] # Объединённые ячейки: [(top_row, left_col, bottom_row, right_col), ...]
        self.max_row = 0
        self.max_column = 0
        self._load_data_from_controller()

    def _load_data_from_controller(self):
        """
        Загружает данные из AppController для указанного листа.
        """
        logger.info(f"Загрузка данных для листа '{self.sheet_name}' через AppController.")
        try:
            # Получаем "сырые" данные из AppController
            raw_data_list, _ = self.app_controller.get_sheet_data(self.sheet_name)
            # logger.debug(f"Получены raw_data: {raw_data_list}")

            # Получаем ID листа для загрузки стилей и объединений
            sheet_id = self._get_sheet_id_by_name(self.sheet_name)
            if sheet_id is None:
                logger.error(f"Не удалось получить ID листа '{self.sheet_name}'.")
                return

            # --- ИСПРАВЛЕНО: Получение и вычисление max_row/max_column ---
            # Получаем метаданные листа из current_project
            sheet_metadata = self.app_controller.current_project.get('sheets', {}).get(self.sheet_name, {})
            stored_max_row = sheet_metadata.get('max_row', 0)
            stored_max_column = sheet_metadata.get('max_column', 0)
            logger.debug(f"Метаданные листа '{self.sheet_name}': stored_max_row={stored_max_row}, stored_max_column={stored_max_column}")

            # Вычисляем максимальные row/col из raw_data_list, чтобы убедиться, что модель охватывает все данные
            calculated_max_row = -1
            calculated_max_column = -1
            if raw_data_list:
                for item in raw_data_list:
                    cell_addr = item.get('cell_address', '')
                    if cell_addr:
                        try:
                            row, col = self._xl_cell_to_row_col(cell_addr)
                            if row > calculated_max_row:
                                calculated_max_row = row
                            if col > calculated_max_column:
                                calculated_max_column = col
                        except ValueError as ve:
                            logger.warning(f"Ошибка преобразования адреса ячейки '{cell_addr}' из raw_data: {ve}")
            logger.debug(f"Вычисленные из raw_data максимальные индексы: calculated_max_row={calculated_max_row}, calculated_max_column={calculated_max_column}")

            # Используем максимум из сохранённых и вычисленных значений
            self.max_row = max(stored_max_row, calculated_max_row)
            self.max_column = max(stored_max_column, calculated_max_column)
            logger.info(f"Окончательные размеры модели для '{self.sheet_name}': max_row={self.max_row}, max_column={self.max_column}")

            # Если данные есть, но max_row/max_column отрицательны (например, raw_data_list пуст или только с ошибками)
            # устанавливаем их в 0, чтобы избежать отрицательных индексов.
            if self.max_row < 0:
                self.max_row = 0
            if self.max_column < 0:
                self.max_column = 0

            # Инициализируем пустую таблицу размером (max_row + 1) x (max_column + 1)
            # +1 потому что max_row/max_column - это индексы (0-based), а размер - количество элементов
            self._data = [[None for _ in range(self.max_column + 1)] for _ in range(self.max_row + 1)]
            self._headers = [self._index_to_column_name(i) for i in range(self.max_column + 1)]
            self._row_headers = [str(i + 1) for i in range(self.max_row + 1)]
            logger.debug(f"Модель инициализирована. Размеры _data: {len(self._data)}x{len(self._data[0]) if self._data else 0}")
            logger.debug(f"Размеры _headers: {len(self._headers)}, _row_headers: {len(self._row_headers)}")

            # Заполняем таблицу данными из raw_data_list
            filled_cells_count = 0
            for item in raw_data_list:
                cell_addr = item.get('cell_address', '')
                value = item.get('value', '')
                # logger.debug(f"Обработка ячейки {cell_addr} со значением {value}")

                if cell_addr:
                    try:
                        row, col = self._xl_cell_to_row_col(cell_addr)
                        # Проверяем, попадает ли ячейка в инициализированные границы
                        # (на случай, если calculated_max_* был больше stored_max_*)
                        if 0 <= row < len(self._data) and 0 <= col < len(self._data[0]):
                            self._data[row][col] = value
                            filled_cells_count += 1
                            # logger.debug(f"Значение {value} записано в ({row}, {col})")
                        else:
                             logger.warning(f"Ячейка {cell_addr} ({row}, {col}) вне инициализированных границ модели ({len(self._data)}, {len(self._data[0])}). Пропущена.")
                    except ValueError as ve:
                        logger.error(f"Ошибка преобразования адреса ячейки '{cell_addr}' при заполнении _data: {ve}")
            logger.info(f"Заполнено {filled_cells_count} ячеек в модели для листа '{self.sheet_name}'.")

            # Загружаем стили и объединения
            self._load_styles_from_controller(sheet_id)
            self._load_merged_cells_from_controller(sheet_id)

            logger.info(f"Данные для листа '{self.sheet_name}' успешно загружены в модель.")

        except Exception as e:
            logger.error(f"Ошибка при загрузке данных для листа '{self.sheet_name}': {e}", exc_info=True)


    def _get_sheet_id_by_name(self, sheet_name: str) -> Optional[int]:
        """
        Получает ID листа по его имени.
        Использует AppController для запроса списка листов и поиска соответствия.
        """
        # AppController.get_sheet_names() возвращает список имён.
        # Нам нужен ID. AppController должен предоставить метод для получения ID по имени.
        # Пока что используем существующий метод get_sheet_names и предположим,
        # что он возвращает список словарей {'name': ..., 'sheet_id': ...}
        # Или мы можем получить все метаданные листов.
        # В AppController есть метод get_sheet_names(), но он возвращает только имена.
        # Попробуем использовать текущий проект.
        sheets_info = self.app_controller.current_project.get('sheets', {})
        for name, info in sheets_info.items():
            if name == sheet_name:
                return info.get('sheet_id') # Предполагаем, что в метаданных есть sheet_id
        # Если не нашли в current_project, можно попробовать запросить из БД напрямую через storage
        # Это менее предпочтительно, так как нарушает принцип централизованного доступа через AppController.
        # Проверим, есть ли у AppController метод для получения sheet_id
        # В AppController есть _get_or_create_sheet_id, но он приватный и создаёт, если нет.
        # Пока оставим как есть и предположим, что sheet_id доступен в current_project.
        # Или, если AppController.current_project не содержит sheet_id, нужно модифицировать AppController
        # или использовать ProjectDBStorage напрямую (что не очень хорошо).
        # Попробуем получить sheet_id через AppController, если он предоставляет такой метод.
        # Если нет, то нужно будет добавить его в AppController или использовать storage.
        # В AppController есть методы, которые принимают sheet_id, например, load_sheet_formulas(sheet_id).
        # Но нет метода для получения sheet_id по имени *из текущего проекта* через публичный интерфейс.
        # Это потенциальная проблема в API AppController.
        # Для обхода: получим список всех листов и их ID через прямой вызов storage, если AppController не предоставляет.
        # logger.warning(f"AppController не предоставляет метод для получения sheet_id по имени. Используем обход.")
        try:
            # AppController.get_sheet_names() возвращает имена. Нужен ID.
            # AppController.project_manager или AppController.data_manager могут помочь, но это внутренности.
            # В AppController есть self.project_manager, который имеет доступ к storage.
            # storage = self.app_controller.project_manager.storage
            # storage.load_all_sheets_metadata(project_id) -> возвращает список {'sheet_id': ..., 'name': ...}
            # Это нарушает инкапсуляцию.
            # Лучше добавить метод в AppController: get_sheet_id_by_name(name)
            # Пока что, если sheet_id не находится в current_project, логируем ошибку.
            # Попробуем использовать AppController.get_sheet_names() и получить ID через storage напрямую.
            # Это наихудший вариант, но может сработать.
            # Получим project_path и создадим временный ProjectDBStorage
            # logger.warning(f"Обходное получение sheet_id для '{sheet_name}' через ProjectDBStorage.")
            # from backend.storage.base import ProjectDBStorage
            # storage = ProjectDBStorage(self.app_controller.project_path)
            # storage.connect()
            # sheets_metadata = storage.load_all_sheets_metadata()
            # storage.disconnect()
            # for sheet_meta in sheets_metadata:
            #     if sheet_meta['name'] == sheet_name:
            #         return sheet_meta['sheet_id']
            # logger.error(f"Не удалось получить sheet_id для '{sheet_name}' через обход.")
            # return None
            # Это не очень красиво. Лучше добавить метод в AppController.
            # Пока оставим как есть и предположим, что sheet_id будет в current_project.
            # Проверим, что возвращает AppController.current_project для 'sheets'.
            # logger.debug(f"current_project['sheets'] = {self.app_controller.current_project.get('sheets', {})}")
            # В текущей реализации AppController.current_project не заполняется информацией о sheet_id.
            # Нужно модифицировать AppController.load_project или AppController.get_sheet_names,
            # чтобы они возвращали ID.
            # В AppController.get_sheet_names() делается запрос к БД: SELECT name FROM sheets.
            # Нужно SELECT name, sheet_id FROM sheets ORDER BY name;
            # И модифицировать возврат функции, чтобы возвращался словарь или список словарей.
            # Это изменение повлияет на API AppController.
            # Пока что, для быстрого решения, я буду использовать обходной путь через storage,
            # но с пониманием, что это временное решение и API AppController стоит улучшить.
            # Попробуем использовать storage через AppController, если возможно.
            storage = self.app_controller.storage
            if storage:
                sheets_metadata = storage.load_all_sheets_metadata()
                for sheet_meta in sheets_metadata:
                    if sheet_meta['name'] == sheet_name:
                        return sheet_meta['sheet_id']
            logger.error(f"Не удалось получить sheet_id для '{sheet_name}' через AppController.storage.")
            return None
        except Exception as e:
            logger.error(f"Ошибка при получении sheet_id: {e}", exc_info=True)
            return None


    def _load_styles_from_controller(self, sheet_id: int):
        """
        Загружает стили ячеек из AppController.
        """
        logger.info(f"Загрузка стилей для листа ID {sheet_id}.")
        try:
            # --- ИСПРАВЛЕНО: Обработка отсутствия метода AppController ---
            # Проверяем, существует ли метод перед вызовом
            if not hasattr(self.app_controller, 'load_sheet_styles'):
                logger.warning(f"AppController не имеет метода 'load_sheet_styles'. Загрузка стилей пропущена для листа ID {sheet_id}.")
                self._styles = {} # Убедимся, что стили пусты
                return

            # Вызываем метод AppController для загрузки стилей для конкретного листа
            # AppController.load_sheet_styles(sheet_id) должен возвращать список
            # [{'range_address': 'A1:B2', 'style_attributes': '{\"font\": {...}, \"fill\": {...}}'}, ...]
            styles_list = self.app_controller.load_sheet_styles(sheet_id)
            logger.debug(f"Получены стили из AppController: {styles_list}")

            # Очищаем старые стили
            self._styles = {}

            # Преобразуем полученные стили в формат self._styles
            for style_item in styles_list:
                range_addr = style_item.get('range_address', '')
                style_attrs_json = style_item.get('style_attributes', '{}')
                if range_addr and style_attrs_json:
                    try:
                        style_attrs = json.loads(style_attrs_json)
                        # logger.debug(f"Десериализованные атрибуты стиля для {range_addr}: {style_attrs}")
                        # Преобразовать range_addr в координаты (row_start, col_start, row_end, col_end)
                        row_start, col_start, row_end, col_end = self._xl_range_to_coords(range_addr)
                        # Заполнить self._styles для каждой ячейки в диапазоне
                        for r in range(row_start, row_end + 1):
                            for c in range(col_start, col_end + 1):
                                # Если ячейка уже имеет стиль, текущая логика перезаписывает его.
                                # В реальных сценариях диапазоны могут пересекаться, и нужно решать, какой стиль приоритетнее.
                                # Для MVP/простоты принимаем стиль из последнего обработанного диапазона.
                                self._styles[(r, c)] = style_attrs
                    except json.JSONDecodeError as je:
                        logger.error(f"Ошибка разбора JSON стиля для диапазона {range_addr}: {je}")
                    except ValueError as ve: # Ошибка от _xl_range_to_coords
                        logger.error(f"Ошибка преобразования диапазона {range_addr}: {ve}")
            logger.info(f"Стили для листа ID {sheet_id} загружены в модель.")

        except AttributeError as ae:
            # Перехватываем конкретное исключение AttributeError, если оно возникло не на hasattr, а при вызове
            logger.error(f"AppController не имеет метода 'load_sheet_styles' (AttributeError): {ae}")
            self._styles = {} # Убедимся, что стили пусты
        except Exception as e:
            logger.error(f"Ошибка при загрузке стилей для листа ID {sheet_id}: {e}", exc_info=True)

    def _load_merged_cells_from_controller(self, sheet_id: int):
        """
        Загружает объединённые ячейки из AppController.
        """
        logger.info(f"Загрузка объединений для листа ID {sheet_id}.")
        try:
            # --- ИСПРАВЛЕНО: Обработка отсутствия метода AppController ---
            # Проверяем, существует ли метод перед вызовом
            if not hasattr(self.app_controller, 'load_sheet_merged_cells'):
                logger.warning(f"AppController не имеет метода 'load_sheet_merged_cells'. Загрузка объединений пропущена для листа ID {sheet_id}.")
                self._merged_cells = [] # Убедимся, что объединения пусты
                return

            # Вызываем метод AppController для загрузки объединений для конкретного листа
            # AppController.load_sheet_merged_cells(sheet_id) должен возвращать список строк ('A1:B2')
            merged_ranges = self.app_controller.load_sheet_merged_cells(sheet_id)
            logger.debug(f"Получены объединения из AppController: {merged_ranges}")

            # Очищаем старые объединения
            self._merged_cells = []

            # Преобразуем полученные строки в координаты
            for range_str in merged_ranges:
                try:
                    row_start, col_start, row_end, col_end = self._xl_range_to_coords(range_str)
                    self._merged_cells.append((row_start, col_start, row_end, col_end))
                except ValueError as ve: # Ошибка от _xl_range_to_coords
                    logger.error(f"Ошибка преобразования диапазона объединения {range_str}: {ve}")
            logger.info(f"Объединения для листа ID {sheet_id} загружены в модель.")

        except AttributeError as ae:
             # Перехватываем конкретное исключение AttributeError
            logger.error(f"AppController не имеет метода 'load_sheet_merged_cells' (AttributeError): {ae}")
            self._merged_cells = [] # Убедимся, что объединения пусты
        except Exception as e:
            logger.error(f"Ошибка при загрузке объединений для листа ID {sheet_id}: {e}", exc_info=True)

    def _xl_cell_to_row_col(self, cell: str) -> tuple[int, int]:
        """
        Преобразует адрес ячейки Excel (e.g., 'A1') в индексы строки и столбца (0-based).
        """
        # Используем логику из xlsxwriter_exporter для согласованности
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
        # logger.debug(f"[КООРД] Результат для '{cell}': {result}")
        return result

    def _index_to_column_name(self, index: int) -> str:
        """
        Преобразует индекс столбца (0-based) в его буквенное обозначение Excel (A, B, ..., Z, AA, AB, ...).
        """
        if index < 0:
            return ""
        result = ""
        while index >= 0:
            result = chr(ord('A') + index % 26) + result
            index = index // 26 - 1
        return result

    def _xl_range_to_coords(self, range_str: str) -> tuple[int, int, int, int]:
        """
        Преобразует диапазон Excel (e.g., 'A1:B10') в координаты (row_start, col_start, row_end, col_end) (0-based).
        """
        # Используем логику из xlsxwriter_exporter для согласованности
        if ':' not in range_str:
            # Это одиночная ячейка
            # logger.debug(f"[КООРД] Диапазон '{range_str}' - это одиночная ячейка.")
            r, c = self._xl_cell_to_row_col(range_str)
            coords = (r, c, r, c)
            # logger.debug(f"[КООРД] Результат для '{range_str}': {coords}")
            return coords

        start_cell, end_cell = range_str.split(':', 1)
        # logger.debug(f"[КООРД] Разделение диапазона на '{start_cell}' и '{end_cell}'.")
        row_start, col_start = self._xl_cell_to_row_col(start_cell)
        row_end, col_end = self._xl_cell_to_row_col(end_cell)
        coords = (row_start, col_start, row_end, col_end)
        # logger.debug(f"[КООРД] Результат для диапазона '{range_str}': {coords}")
        return coords

    def rowCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = QModelIndex()) -> int: # <-- Изменено
        """Возвращает количество строк."""
        if parent.isValid():
            return 0
        return len(self._data)

    def columnCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = QModelIndex()) -> int: # <-- Изменено
        """Возвращает количество столбцов."""
        if parent.isValid():
            return 0
        return len(self._data[0]) if self._data else 0

    def data(self, index: Union[QModelIndex, QPersistentModelIndex], role: int = Qt.ItemDataRole.DisplayRole) -> Any: # <-- Изменено
        """Возвращает данные для указанной ячейки и роли."""
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if row >= len(self._data) or col >= len(self._data[0]):
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._data[row][col]
            # Возвращаем строковое представление значения
            return str(value) if value is not None else ""
        elif role == Qt.ItemDataRole.EditRole:
            # Для редактирования возвращаем "сырое" значение
            return self._data[row][col]
        elif role == Qt.ItemDataRole.BackgroundRole:
            # Вернуть QBrush для фона ячейки на основе стиля
            style = self._styles.get((row, col), {})
            bg_color_hex = style.get('bg_color')
            if bg_color_hex:
                return QBrush(QColor(f"#{bg_color_hex}"))
        elif role == Qt.ItemDataRole.ForegroundRole:
            # Вернуть QBrush для текста ячейки на основе стиля
            style = self._styles.get((row, col), {})
            font_color_hex = style.get('font_color')
            if font_color_hex:
                return QBrush(QColor(f"#{font_color_hex}"))
        elif role == Qt.ItemDataRole.FontRole:
            # Вернуть QFont для ячейки на основе стиля
            style = self._styles.get((row, col), {})
            font_attrs = {
                'bold': style.get('bold'),
                'italic': style.get('italic'),
                'underline': style.get('underline'),
                'font_size': style.get('font_size'),
                'font_name': style.get('font_name')
            }
            # Создаём QFont и применяем атрибуты
            font = QFont()
            if font_attrs['font_name']:
                font.setFamily(font_attrs['font_name'])
            if font_attrs['font_size']:
                font.setPointSize(int(font_attrs['font_size']))
            if font_attrs['bold'] is not None:
                font.setBold(bool(font_attrs['bold']))
            if font_attrs['italic'] is not None:
                font.setItalic(bool(font_attrs['italic']))
            # underline может быть строкой ('single', 'double', ...), обработка требует уточнения
            # if font_attrs['underline']:
            #     font.setUnderline(True) # Упрощённо
            return font
        elif role == Qt.ItemDataRole.TextAlignmentRole:
            # Вернуть флаги выравнивания на основе стиля
            style = self._styles.get((row, col), {})
            align_h = style.get('align', 'left') # 'left', 'center', 'right', 'fill', 'justify', 'distributed'
            align_v = style.get('valign', 'top') # 'top', 'vcenter', 'bottom', 'vjustify', 'vdistributed'
            # Сопоставление строк с флагами Qt
            align_map_h = {
                'left': Qt.AlignmentFlag.AlignLeft,
                'center': Qt.AlignmentFlag.AlignHCenter,
                'right': Qt.AlignmentFlag.AlignRight,
                'fill': Qt.AlignmentFlag.AlignJustify,
                'justify': Qt.AlignmentFlag.AlignJustify,
                'distributed': Qt.AlignmentFlag.AlignJustify,
            }
            align_map_v = {
                'top': Qt.AlignmentFlag.AlignTop,
                'vcenter': Qt.AlignmentFlag.AlignVCenter,
                'bottom': Qt.AlignmentFlag.AlignBottom,
                'vjustify': Qt.AlignmentFlag.AlignJustify,
                'vdistributed': Qt.AlignmentFlag.AlignJustify,
            }
            h_flag = align_map_h.get(align_h, Qt.AlignmentFlag.AlignLeft)
            v_flag = align_map_v.get(align_v, Qt.AlignmentFlag.AlignTop)
            return h_flag | v_flag

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        """Возвращает данные для заголовков строк или столбцов."""
        if role != Qt.ItemDataRole.DisplayRole:
            return None

        if orientation == Qt.Orientation.Horizontal:
            if 0 <= section < len(self._headers):
                return self._headers[section]
        elif orientation == Qt.Orientation.Vertical:
            if 0 <= section < len(self._row_headers):
                return self._row_headers[section]
        return None

    def setData(self, index: Union[QModelIndex, QPersistentModelIndex], value: Any, role: int = Qt.ItemDataRole.EditRole) -> bool: # <-- Изменено
        """
        Устанавливает данные для указанной ячейки.
        Вызывается при редактировании в QTableView.
        """
        # TODO: Реализовать изменение данных через AppController
        # 1. Проверить, что index.isValid() и role == Qt.EditRole
        # 2. Преобразовать index.row(), index.column() в адрес ячейки Excel (e.g., 'A1')
        # 3. Вызвать метод AppController для обновления значения ячейки
        #    app_controller.update_cell_value(self.sheet_name, cell_address, value)
        # 4. Обновить внутреннее состояние модели (self._data)
        # 5. Вызвать self.dataChanged.emit() для уведомления представления
        #    self.dataChanged.emit(index, index, [role])
        # logger.debug(f"Попытка установки данных в ячейку ({index.row()}, {index.column()}) значение {value}")
        if not index.isValid() or role != Qt.ItemDataRole.EditRole:
            return False

        row = index.row()
        col = index.column()
        if row >= len(self._data) or col >= len(self._data[0]):
            return False

        cell_address = self._index_to_cell_address(row, col)
        try:
            # Вызов AppController для обновления ячейки
            success = self.app_controller.update_cell_value(self.sheet_name, cell_address, value)
            if success:
                # Обновляем локальное состояние модели
                self._data[row][col] = value
                # Уведомляем представление об изменении
                self.dataChanged.emit(index, index, [role])
                logger.info(f"Ячейка {cell_address} обновлена через AppController.")
                return True
            else:
                logger.error(f"AppController не смог обновить ячейку {cell_address}.")
                return False
        except Exception as e:
            logger.error(f"Ошибка при обновлении ячейки {cell_address} через AppController: {e}", exc_info=True)
            return False

    def flags(self, index: Union[QModelIndex, QPersistentModelIndex]) -> Qt.ItemFlag: # <-- Изменено
        """Определяет флаги для ячейки (редактируемая, доступная и т.д.)."""
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags # <-- Исправлено ItemFlags -> ItemFlag

        # Делаем ячейки редактируемыми
        return super().flags(index) | Qt.ItemFlag.ItemIsEditable # <-- Исправлено ItemFlags -> ItemFlag

    # --- Новый метод для вставки данных из буфера обмена ---
    def insert_data_from_clipboard(self, parsed_data: List[List[str]], start_row: int, start_col: int):
        """
        Вставляет данные из буфера обмена в модель и БД, начиная с указанной ячейки.

        Args:
            parsed_data (List[List[str]]): Двумерный список значений для вставки.
            start_row (int): Начальная строка (0-based).
            start_col (int): Начальный столбец (0-based).
        """
        if not parsed_data or not parsed_data[0]:
            logger.debug("Попытка вставить пустые данные из буфера.")
            return

        logger.info(f"Начало вставки данных из буфера. Размер: {len(parsed_data)}x{len(parsed_data[0])}. Начало: ({start_row}, {start_col})")

        # Определим новые максимальные размеры таблицы после вставки
        end_row = start_row + len(parsed_data) - 1
        end_col = start_col + len(parsed_data[0]) - 1
        new_max_row = max(self.max_row, end_row)
        new_max_column = max(self.max_column, end_col)

        # --- Пакетная вставка через AppController ---
        # Соберём все изменения в один список для потенциальной оптимизации в AppController
        # или выполним в цикле, если AppController не поддерживает пакетную вставку напрямую
        # через update_cell_value.
        # Пока что используем update_cell_value в цикле.
        # TODO: Рассмотреть добавление метода в AppController для пакетного обновления ячеек.
        batch_success = True
        for r_idx, row_data in enumerate(parsed_data):
            for c_idx, cell_value in enumerate(row_data):
                abs_row = start_row + r_idx
                abs_col = start_col + c_idx
                cell_address = self._index_to_cell_address(abs_row, abs_col)

                try:
                    # Вызываем AppController для обновления ячейки в БД
                    success = self.app_controller.update_cell_value(self.sheet_name, cell_address, cell_value)
                    if not success:
                        logger.error(f"AppController не смог обновить ячейку {cell_address} при вставке из буфера.")
                        batch_success = False
                        # В реальной реализации возможно прерывание или откат транзакции
                        # Пока просто логируем ошибку и продолжаем
                except Exception as e:
                    logger.error(f"Ошибка при обновлении ячейки {cell_address} через AppController: {e}", exc_info=True)
                    batch_success = False
                    # В реальной реализации возможно прерывание или откат транзакции
                    # Пока просто логируем ошибку и продолжаем

        if not batch_success:
            logger.error("Одна или несколько ячеек не были обновлены при вставке из буфера.")
            # В реальной реализации можно показать сообщение пользователю
            return

        # --- Обновление внутреннего представления модели ---
        # Увеличиваем размеры _data, если нужно
        if new_max_row > self.max_row or new_max_column > self.max_column:
            # Увеличиваем количество строк
            current_row_count = len(self._data)
            if new_max_row + 1 > current_row_count:
                for _ in range(new_max_row + 1 - current_row_count):
                    self._data.append([None for _ in range(self.max_column + 1)])
            # Увеличиваем количество столбцов в каждой строке
            current_col_count = len(self._data[0]) if self._data else 0
            if new_max_column + 1 > current_col_count:
                for row in self._data:
                    row.extend([None for _ in range(new_max_column + 1 - current_col_count)])
            # Обновляем max_ переменные
            self.max_row = new_max_row
            self.max_column = new_max_column
            # Обновляем заголовки
            self._headers = [self._index_to_column_name(i) for i in range(self.max_column + 1)]
            self._row_headers = [str(i + 1) for i in range(self.max_row + 1)]
            logger.info(f"Размеры модели увеличены до ({self.max_row + 1}, {self.max_column + 1}).")

        # Заполняем _data новыми значениями
        for r_idx, row_data in enumerate(parsed_data):
            for c_idx, cell_value in enumerate(row_data):
                abs_row = start_row + r_idx
                abs_col = start_col + c_idx
                if 0 <= abs_row < len(self._data) and 0 <= abs_col < len(self._data[0]):
                    self._data[abs_row][abs_col] = cell_value
                    logger.debug(f"Значение '{cell_value}' вставлено в модель в ячейку ({abs_row}, {abs_col})")

        # Уведомляем QTableView об изменениях
        # layoutChanged.emit() говорит представлению, что структура данных (размеры) могла измениться.
        # dataChanged.emit() говорит, что содержимое изменилось.
        # Так как мы могли изменить размеры и содержимое, вызовем оба.
        # top_left = self.index(start_row, start_col)
        # bottom_right = self.index(min(end_row, self.max_row), min(end_col, self.max_column))
        # self.dataChanged.emit(top_left, bottom_right, [Qt.DisplayRole])
        # self.layoutChanged.emit() # Это может быть избыточно, если мы точно знаем, что изменили только данные.
        # Более точный способ - уведомить только об изменении данных, если размеры не менялись.
        # Но так как размеры *могут* измениться, и мы обновляем _data/_headers/_row_headers,
        # layoutChanged.emit() более безопасен, чтобы QTableView пересчитал размеры.
        # Однако, layoutChanged может быть ресурсоемким. Попробуем с dataChanged и проверим.
        # Если вставка за пределы текущего размера не отображается сразу, добавим layoutChanged.
        # Для начала вызовем dataChanged для диапазона вставки.
        top_left = self.index(start_row, start_col)
        bottom_right = self.index(min(end_row, self.max_row), min(end_col, self.max_column))
        self.dataChanged.emit(top_left, bottom_right, [Qt.ItemDataRole.DisplayRole])

        # Если размеры таблицы изменились, вызываем layoutChanged.
        if new_max_row > self.max_row or new_max_column > self.max_column:
             self.layoutChanged.emit()
        # Или, если мы точно уверены, что размеры могли измениться, всегда вызываем layoutChanged.
        # self.layoutChanged.emit() # Это более консервативный подход.

        logger.info(f"Вставка данных из буфера завершена. Обновлён диапазон ({start_row}, {start_col}) - ({end_row}, {end_col}).")

    # --- Конец нового метода ---

    def _index_to_cell_address(self, row: int, col: int) -> str:
        """
        Преобразует индексы строки и столбца (0-based) в адрес ячейки Excel (e.g., 'A1').
        """
        col_name = self._index_to_column_name(col)
        row_num = row + 1
        return f"{col_name}{row_num}"

    # --- Исправленный метод span ---
    def span(self, index: Union[QModelIndex, QPersistentModelIndex]) -> QSize: # <-- Изменена сигнатура
        """
        Возвращает размер объединённой области для ячейки index.
        Используется QTableView для отображения объединённых ячеек.
        """
        if not index.isValid():
            return QSize(1, 1)

        row = index.row()
        col = index.column()

        # logger.debug(f"Запрос span для ({row}, {column})")
        for top_row, left_col, bottom_row, right_col in self._merged_cells:
            if top_row <= row <= bottom_row and left_col <= col <= right_col:
                row_span = bottom_row - top_row + 1
                col_span = right_col - left_col + 1
                # logger.debug(f"Ячейка ({row}, {column}) принадлежит объединению ({top_row}, {left_col})-({bottom_row}, {right_col}). Span: ({row_span}, {col_span})")
                return QSize(col_span, row_span) # <-- Возвращаем QSize(width, height), где width - колонки, height - строки
        # logger.debug(f"Ячейка ({row}, {column}) не объединена. Span: (1, 1)")
        return QSize(1, 1) # <-- Возвращаем QSize(1, 1)

    # --- Конец исправленного метода span ---
