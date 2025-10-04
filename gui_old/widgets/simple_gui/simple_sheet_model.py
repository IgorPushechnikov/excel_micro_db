# backend/constructor/widgets/simple_gui/simple_sheet_model.py
"""
Модель данных для упрощённого табличного редактора.
Загружает данные и стили из AppController и предоставляет их QTableView.
"""
import json
import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QPersistentModelIndex, Qt
from PySide6.QtGui import QBrush, QColor, QFont

# УДАЛЯЕМ: from backend.constructor.widgets.simple_gui.db_data_fetcher import fetch_sheet_data

# Импортируем вспомогательные функции для парсинга адресов
from backend.constructor.widgets.simple_gui.db_data_fetcher import _parse_cell_address, _parse_range_address


class SimpleSheetModel(QAbstractTableModel):
    """Модель данных для отображения содержимого листа Excel."""

    def __init__(self, raw_data_list, styles_list, parent=None):
        """
        Инициализирует модель с данными, полученными от AppController.

        Args:
            raw_data_list (List[Dict[str, Any]]): Список словарей с 'cell_address' и 'value'.
            styles_list (List[Dict[str, Any]]): Список словарей с 'range_address' и 'style_attributes'.
            parent: Родительский объект.
        """
        super().__init__(parent)
        # УДАЛЯЕМ: self.db_path = db_path
        # УДАЛЯЕМ: self.sheet_name = sheet_name

        # Настройка логгера (упрощённая версия, без привязки к БД)
        self.logger = logging.getLogger(__name__ + ".SimpleSheetModel")

        # Данные таблицы
        self._data: List[List[Any]] = []
        # Стили ячеек: {(row, col): {"font_color": "#FF0000", "bg_color": "#FFFFFF", ...}}
        self._cell_styles: Dict[Tuple[int, int], Dict[str, Any]] = {}

        # Заголовки столбцов Excel (A, B, C...)
        self._column_headers: List[str] = []

        # Максимальные индексы для rowCount/columnCount
        self._max_row = -1
        self._max_col = -1

        self.logger.info(f"Инициализация SimpleSheetModel с переданными данными")
        self._process_raw_data(raw_data_list)
        self._process_styles(styles_list)
        self._generate_column_headers()
        self.logger.info(f"SimpleSheetModel успешно инициализирована, max_row={self._max_row}, max_col={self._max_col}.")



    def _process_raw_data(self, raw_data_list):
        """Обрабатывает raw_data_list от AppController в 2D список _data."""
        try:
            max_row = -1
            max_col = -1
            data_map = {}

            for item in raw_data_list:
                addr = item.get('cell_address')
                val = item.get('value')
                parsed = _parse_cell_address(addr)
                if parsed is None:
                    self.logger.warning(f"Не удалось распарсить адрес ячейки '{addr}'")
                    continue
                row_idx, col_idx = parsed
                data_map[(row_idx, col_idx)] = val
                max_row = max(max_row, row_idx)
                max_col = max(max_col, col_idx)

            # Создание 2D списка
            if max_row >= 0 and max_col >= 0:
                self._data = [[""] * (max_col + 1) for _ in range(max_row + 1)]
                for (r, c), val in data_map.items():
                    if 0 <= r <= max_row and 0 <= c <= max_col:
                        self._data[r][c] = val
            else:
                self._data = []

            self._max_row = max_row
            self._max_col = max_col

            self.logger.info(f"Данные обработаны: max_row={self._max_row}, max_col={self._max_col}")

        except Exception as e:
            self.logger.error(f"Ошибка при обработке raw_data_list: {e}", exc_info=True)
            # Инициализируем пустую модель в случае ошибки
            self._data = []
            self._max_row = -1
            self._max_col = -1

    def _process_styles(self, styles_list):
        """Обрабатывает styles_list от AppController в словарь _cell_styles."""
        try:
            for item in styles_list:
                range_addr = item.get('range_address')
                style_attrs_json = item.get('style_attributes')
                parsed_range = _parse_range_address(range_addr)
                if parsed_range is None:
                    self.logger.warning(f"Не удалось распарсить адрес диапазона '{range_addr}'")
                    continue
                start_r, start_c, end_r, end_c = parsed_range

                try:
                    style_attrs = json.loads(style_attrs_json) if isinstance(style_attrs_json, str) else style_attrs_json
                    if not isinstance(style_attrs, dict):
                         self.logger.warning(f"'style_attributes' для '{range_addr}' не является словарем или корректным JSON-объектом. Пропущено.")
                         continue
                except json.JSONDecodeError as e:
                    self.logger.warning(f"Ошибка парсинга JSON стиля для '{range_addr}': {e}")
                    continue # Пропускаем некорректный стиль

                # Применяем стиль ко всем ячейкам в диапазоне
                for r in range(start_r, end_r + 1):
                    for c in range(start_c, end_c + 1):
                        self._cell_styles[(r, c)] = style_attrs

            self.logger.info(f"Стили обработаны: {len(self._cell_styles)} стилей для ячеек")

        except Exception as e:
            self.logger.error(f"Ошибка при обработке styles_list: {e}", exc_info=True)
            # Инициализируем пустой словарь стилей в случае ошибки
            self._cell_styles = {}

    def _generate_column_headers(self):
        """Генерирует имена столбцов Excel (A, B, ..., Z, AA, AB...)."""
        self._column_headers = []
        if self._max_col < 0:
            return

        for i in range(self._max_col + 1):
            name = ""
            temp = i
            while True:
                name = chr(temp % 26 + ord('A')) + name
                temp = temp // 26 - 1
                if temp < 0:
                    break
            self._column_headers.append(name)

    def _column_index_to_letter(self, index: int) -> str:
        """Преобразует 0-базовый индекс столбца в букву Excel (A, B, ..., AA, AB...)."""
        if index < 0:
            return ""
        name = ""
        temp = index
        while True:
            name = chr(temp % 26 + ord('A')) + name
            temp = temp // 26 - 1
            if temp < 0:
                break
        return name

    # Реализация QAbstractTableModel
    def rowCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return self._max_row + 1 if self._max_row >= 0 else 0

    def columnCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return self._max_col + 1 if self._max_col >= 0 else 0

    def data(self, index: QModelIndex | QPersistentModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid():
            self.logger.debug("data: index is not valid")
            return None

        row = index.row()
        col = index.column()

        if row > self._max_row or col > self._max_col or row < 0 or col < 0:
            self.logger.debug(f"data: index out of bounds: row={row}, col={col}, max_row={self._max_row}, max_col={self._max_col}")
            return None

        try:
            if role == Qt.ItemDataRole.DisplayRole:
                value = self._data[row][col]
                self.logger.debug(f"data: DisplayRole for [{row}, {col}] = '{value}'")
                return value
            elif role == Qt.ItemDataRole.BackgroundRole:
                style = self._cell_styles.get((row, col))
                if style:
                    bg_color = style.get("fill_fg_color") or style.get("fill_bg_color")
                    if bg_color and isinstance(bg_color, str) and bg_color.startswith('#'):
                        self.logger.debug(f"data: BackgroundRole for [{row}, {col}] = '{bg_color}'")
                        return QBrush(QColor(bg_color))
            elif role == Qt.ItemDataRole.ForegroundRole:
                style = self._cell_styles.get((row, col))
                if style:
                    font_color = style.get("font_color")
                    if font_color and isinstance(font_color, str) and font_color.startswith('#'):
                        self.logger.debug(f"data: ForegroundRole for [{row}, {col}] = '{font_color}'")
                        return QBrush(QColor(font_color))
            elif role == Qt.ItemDataRole.FontRole:
                style = self._cell_styles.get((row, col))
                if style:
                    font = QFont()
                    # Жирность
                    font_b = style.get("font_b")
                    if font_b in (True, '1', 'true', 'True'):
                        font.setBold(True)
                    # Курсив
                    font_i = style.get("font_i")
                    if font_i in (True, '1', 'true', 'True'):
                        font.setItalic(True)
                    # Подчеркивание
                    font_u = style.get("font_u")
                    if font_u and font_u != "none":
                        font.setUnderline(True)
                    # Зачеркивание
                    font_strike = style.get("font_strike")
                    if font_strike in (True, '1', 'true', 'True'):
                        font.setStrikeOut(True)
                    self.logger.debug(f"data: FontRole for [{row}, {col}] applied")
                    return font
        except Exception as e:
            self.logger.error(f"data: Ошибка при обработке роли {role} для [{row}, {col}]: {e}", exc_info=True)

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role == Qt.ItemDataRole.DisplayRole:
            try:
                if orientation == Qt.Orientation.Horizontal:
                    if 0 <= section <= self._max_col:
                        header = self._column_headers[section]
                        self.logger.debug(f"headerData: Horizontal header for section {section} = '{header}'")
                        return header
                elif orientation == Qt.Orientation.Vertical:
                    header = str(section + 1)  # Номера строк 1-based
                    self.logger.debug(f"headerData: Vertical header for section {section} = '{header}'")
                    return header
            except Exception as e:
                self.logger.error(f"headerData: Ошибка при получении заголовка для section {section}, orientation {orientation}: {e}", exc_info=True)
        return None

    def flags(self, index: QModelIndex | QPersistentModelIndex) -> Qt.ItemFlag:
        """Возвращает флаги для ячейки. Ячейки только для чтения."""
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable