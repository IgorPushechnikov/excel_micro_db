# backend/constructor/widgets/simple_gui/simple_sheet_model.py
"""
Модель данных для упрощённого табличного редактора.
Загружает данные и стили из БД через db_data_fetcher и предоставляет их QTableView.
"""
import json
import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor, QFont

from backend.constructor.widgets.simple_gui.db_data_fetcher import fetch_sheet_data


class SimpleSheetModel(QAbstractTableModel):
    """Модель данных для отображения содержимого листа Excel."""

    def __init__(self, db_path: str, sheet_name: str, parent=None):
        """
        Инициализирует модель с данными из БД.

        Args:
            db_path (str): Путь к файлу БД.
            sheet_name (str): Имя листа для загрузки.
            parent: Родительский объект.
        """
        super().__init__(parent)
        self.db_path = db_path
        self.sheet_name = sheet_name

        # Настройка логгера для записи в файл проекта
        self.logger = self._setup_logger()

        # Данные таблицы
        self._data: List[List[Any]] = []
        # Стили ячеек: {(row, col): {"font_color": "#FF0000", "bg_color": "#FFFFFF", ...}}
        self._cell_styles: Dict[Tuple[int, int], Dict[str, Any]] = {}

        # Заголовки столбцов Excel (A, B, C...)
        self._column_headers: List[str] = []

        # Максимальные индексы для rowCount/columnCount
        self._max_row = -1
        self._max_col = -1

        self.logger.info(f"Инициализация SimpleSheetModel для листа '{self.sheet_name}' из БД: {self.db_path}")
        self._load_data_from_fetcher()
        self._generate_column_headers()
        self.logger.info(f"SimpleSheetModel для листа '{self.sheet_name}' успешно инициализирована.")

    def _setup_logger(self) -> logging.Logger:
        """Настраивает логгер для записи в файл проекта."""
        logger_name = f"SimpleSheetModel_{self.sheet_name}"
        logger = logging.getLogger(logger_name)
        logger.setLevel(logging.DEBUG)

        # Очищаем существующие хендлеры, чтобы избежать дублирования
        logger.handlers.clear()

        # Определяем путь к папке проекта и папке логов
        db_path_obj = Path(self.db_path)
        project_dir = db_path_obj.parent
        logs_dir = project_dir / "logs"
        logs_dir.mkdir(exist_ok=True)  # Создаем папку logs, если она не существует

        log_file_path = logs_dir / "simple_sheet_model.log"

        # Создаем FileHandler
        file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)

        # Форматтер для логов
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        return logger

    def _load_data_from_fetcher(self):
        """Загружает данные и стили из БД через db_data_fetcher."""
        try:
            self.logger.debug("Вызов db_data_fetcher.fetch_sheet_data...")
            rows_2d, styles_map = fetch_sheet_data(self.sheet_name, self.db_path)
            self.logger.debug(f"Получено от fetch_sheet_data: {len(rows_2d)} строк, {len(styles_map)} стилей")

            self._data = rows_2d
            self._cell_styles = styles_map

            # Вычисляем max_row и max_col
            if self._data:
                self._max_row = len(self._data) - 1
                self._max_col = len(self._data[0]) - 1 if self._data[0] else -1
            else:
                self._max_row = -1
                self._max_col = -1

            self.logger.info(f"Данные загружены: max_row={self._max_row}, max_col={self._max_col}")

        except Exception as e:
            self.logger.error(f"Ошибка при загрузке данных через fetch_sheet_data: {e}", exc_info=True)
            # Инициализируем пустую модель в случае ошибки
            self._data = []
            self._cell_styles = {}
            self._max_row = -1
            self._max_col = -1

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
    def rowCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return self._max_row + 1 if self._max_row >= 0 else 0

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return self._max_col + 1 if self._max_col >= 0 else 0

    def data(self, index: QModelIndex, role=Qt.ItemDataRole.DisplayRole) -> Any:
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

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.ItemDataRole.DisplayRole) -> Any:
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

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        """Возвращает флаги для ячейки. Ячейки только для чтения."""
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable