# backend/constructor/widgets/new_gui/qt_model_adapter.py
"""
Модуль-адаптер для связи между AppController/БД и QAbstractTableModel (PySide6).
Предоставляет данные из БД для отображения в QTableView.
"""

import json
import logging
from typing import Any, Dict, List, Optional
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
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

            # Получаем метаданные листа
            sheet_metadata = self.app_controller.current_project.get('sheets', {}).get(self.sheet_name, {})
            self.max_row = sheet_metadata.get('max_row', 0)
            self.max_column = sheet_metadata.get('max_column', 0)

            # Инициализируем пустую таблицу размером max_row x max_col
            self._data = [[None for _ in range(self.max_column + 1)] for _ in range(self.max_row + 1)]
            self._headers = [self._index_to_column_name(i) for i in range(self.max_column + 1)]
            self._row_headers = [str(i + 1) for i in range(self.max_row + 1)]

            # Заполняем таблицу данными из raw_data_list
            for item in raw_data_list:
                cell_addr = item.get('cell_address', '')
                value = item.get('value', '')
                # logger.debug(f"Обработка ячейки {cell_addr} со значением {value}")

                if cell_addr:
                    row, col = self._xl_cell_to_row_col(cell_addr)
                    if 0 <= row <= self.max_row and 0 <= col <= self.max_column:
                        self._data[row][col] = value
                        # logger.debug(f"Значение {value} записано в ({row}, {col})")

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

        except Exception as e:
            logger.error(f"Ошибка при загрузке стилей для листа ID {sheet_id}: {e}", exc_info=True)

    def _load_merged_cells_from_controller(self, sheet_id: int):
        """
        Загружает объединённые ячейки из AppController.
        """
        logger.info(f"Загрузка объединений для листа ID {sheet_id}.")
        try:
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

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Возвращает количество строк."""
        if parent.isValid():
            return 0
        return len(self._data)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Возвращает количество столбцов."""
        if parent.isValid():
            return 0
        return len(self._data[0]) if self._data else 0

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
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

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.ItemDataRole.EditRole) -> bool:
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

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        """Определяет флаги для ячейки (редактируемая, доступная и т.д.)."""
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags

        # Делаем ячейки редактируемыми
        return super().flags(index) | Qt.ItemFlag.ItemIsEditable

    def _index_to_cell_address(self, row: int, col: int) -> str:
        """
        Преобразует индексы строки и столбца (0-based) в адрес ячейки Excel (e.g., 'A1').
        """
        col_name = self._index_to_column_name(col)
        row_num = row + 1
        return f"{col_name}{row_num}"

    def span(self, row: int, column: int) -> tuple[int, int]:
        """
        Возвращает размер объединённой области для ячейки (row, column).
        Используется QTableView для отображения объединённых ячеек.
        """
        # logger.debug(f"Запрос span для ({row}, {column})")
        for top_row, left_col, bottom_row, right_col in self._merged_cells:
            if top_row <= row <= bottom_row and left_col <= column <= right_col:
                row_span = bottom_row - top_row + 1
                col_span = right_col - left_col + 1
                # logger.debug(f"Ячейка ({row}, {column}) принадлежит объединению ({top_row}, {left_col})-({bottom_row}, {right_col}). Span: ({row_span}, {col_span})")
                return row_span, col_span
        # logger.debug(f"Ячейка ({row}, {column}) не объединена. Span: (1, 1)")
        return 1, 1
