# backend/constructor/widgets/new_gui/table_model.py
"""
Модель данных для QTableView в новом GUI.
Предоставляет данные из AppController для отображения в таблице.
"""

import logging
from typing import Any, Dict, List, Optional, Union
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex, QPersistentModelIndex, QSize
from PySide6.QtGui import QFont, QColor, QBrush, QTextOption

# Импортируем AppController
from backend.core.app_controller import create_app_controller
from backend.utils.logger import get_logger

logger = get_logger(__name__)

# --- Константы для модели ---
# Максимальное количество строк/столбцов для отображения в модели по умолчанию
# В реальном приложении это будет определяться метаданными листа
DEFAULT_MAX_ROWS = 1048576  # Excel 2007+
DEFAULT_MAX_COLS = 16384    # Excel 2007+

# Минимальное количество строк/столбцов для отображения, если метаданные не доступны
MIN_DISPLAY_ROWS = 100
MIN_DISPLAY_COLS = 26 # A-Z

# --- Вспомогательные функции для работы с адресами ячеек ---
def _index_to_column_name(index: int) -> str:
    """
    Преобразует индекс столбца (0-based) в его буквенное обозначение Excel (A, B, ..., Z, AA, AB, ...).
    """
    if index < 0:
        return ""
    result = ""
    temp_index = index
    while temp_index >= 0:
        result = chr(ord('A') + temp_index % 26) + result
        temp_index = temp_index // 26 - 1
    return result

def _xl_cell_to_row_col(cell: str) -> tuple[int, int]:
    """
    Преобразует адрес ячейки Excel (e.g., 'A1') в индексы строки и столбца (0-based).
    """
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
# --- Конец вспомогательных функций ---

class TableModel(QAbstractTableModel):
    """
    Модель данных для QTableView, получающая данные из AppController/БД.
    """

    def __init__(self, app_controller, parent=None):
        """
        Инициализирует модель.

        Args:
            app_controller: Экземпляр AppController.
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self.app_controller = app_controller
        self._sheet_name: Optional[str] = None
        
        # Данные листа
        self._display_data: List[List[Any]] = []  # Список списков: строки x колонки
        self._headers: List[str] = []              # Заголовки колонок (A, B, C...)
        self._row_headers: List[str] = []          # Заголовки строк (1, 2, 3...)
        
        # Метаданные листа
        self._max_row = 0
        self._max_column = 0
        
        # Стили (пока без реализации)
        self._styles: Dict[tuple, Dict[str, Any]] = {} # Стили ячеек: {(row, col): {'font': ..., 'bg_color': ...}}
        
        logger.debug("TableModel инициализирована.")

    def load_sheet(self, sheet_name: str):
        """
        Загружает данные из AppController для указанного листа.

        Args:
            sheet_name (str): Имя листа для отображения.
        """
        logger.info(f"Загрузка данных для листа '{sheet_name}' через AppController.")
        try:
            if not self.app_controller or not self.app_controller.is_project_loaded:
                logger.error("Проект не загружен в AppController.")
                self._clear_data()
                self._sheet_name = None
                self.modelReset.emit() # Уведомить представление о сбросе модели
                return

            # Получаем "сырые" данные и редактируемые данные из AppController
            # raw_data_list - это [{'cell_address': 'A1', 'value': '...', 'value_type': '...'}, ...]
            # editable_data - это {'data': [...], 'styles': [...], 'formulas': [...]}
            
            # raw_data_list, editable_data = self.app_controller.get_sheet_data(sheet_name)
            # Для упрощения, пока используем только raw_data_list
            # editable_data может быть использован позже для стилей и формул
            
            # Получаем только "сырые" данные
            raw_data_list = self.app_controller.get_sheet_raw_data(sheet_name)
            
            if raw_data_list is None:
                logger.warning(f"Не удалось получить данные для листа '{sheet_name}'.")
                self._clear_data()
                self._sheet_name = sheet_name
                self.modelReset.emit()
                return
                
            logger.debug(f"Получены raw_data для листа '{sheet_name}': {len(raw_data_list) if raw_data_list else 0} записей.")

            # --- Получение метаданных листа ---
            # Предполагаем, что AppController предоставляет метод для получения метаданных
            # Например, через self.app_controller.storage.load_sheet_metadata(sheet_name)
            # Или через специальный метод в AppController
            # Пока используем заглушку или получим из raw_data
            
            # Вычисляем максимальные row/col из raw_data_list
            calculated_max_row = -1
            calculated_max_column = -1
            if raw_data_list:
                for item in raw_data_list:
                    cell_addr = item.get('cell_address', '')
                    if cell_addr:
                        try:
                            row, col = _xl_cell_to_row_col(cell_addr)
                            if row > calculated_max_row:
                                calculated_max_row = row
                            if col > calculated_max_column:
                                calculated_max_column = col
                        except ValueError as ve:
                            logger.warning(f"Ошибка преобразования адреса ячейки '{cell_addr}' из raw_data: {ve}")

            logger.debug(f"Вычисленные из raw_data максимальные индексы: calculated_max_row={calculated_max_row}, calculated_max_column={calculated_max_column}")

            # Используем вычисленные значения или минимальные значения по умолчанию
            self._max_row = max(calculated_max_row, MIN_DISPLAY_ROWS - 1)
            self._max_column = max(calculated_max_column, MIN_DISPLAY_COLS - 1)
            
            logger.info(f"Окончательные размеры модели для '{sheet_name}': max_row={self._max_row}, max_column={self._max_column}")

            # --- Инициализация структуры данных ---
            # Создаем пустую таблицу размером (max_row + 1) x (max_column + 1)
            # +1 потому что max_row/max_column - это индексы (0-based), а размер - количество элементов
            self.beginResetModel() # Начинаем сброс модели
            try:
                self._display_data = [[None for _ in range(self._max_column + 1)] for _ in range(self._max_row + 1)]
                self._headers = [_index_to_column_name(i) for i in range(self._max_column + 1)]
                self._row_headers = [str(i + 1) for i in range(self._max_row + 1)]
                
                logger.debug(f"Модель инициализирована. Размеры _display_data: {len(self._display_data)}x{len(self._display_data[0]) if self._display_data else 0}")
                logger.debug(f"Размеры _headers: {len(self._headers)}, _row_headers: {len(self._row_headers)}")

                # --- Заполнение таблицы данными из raw_data_list ---
                filled_cells_count = 0
                for item in raw_data_list:
                    cell_addr = item.get('cell_address', '')
                    value = item.get('value', '')
                    
                    if cell_addr:
                        try:
                            row, col = _xl_cell_to_row_col(cell_addr)
                            # Проверяем, попадает ли ячейка в инициализированные границы
                            if 0 <= row < len(self._display_data) and 0 <= col < len(self._display_data[0]):
                                self._display_data[row][col] = value
                                filled_cells_count += 1
                                logger.debug(f"Значение {value} записано в ({row}, {col})")
                            else:
                                 logger.warning(f"Ячейка {cell_addr} ({row}, {col}) вне инициализированных границ модели ({len(self._display_data)}, {len(self._display_data[0])}). Пропущена.")
                        except ValueError as ve:
                            logger.error(f"Ошибка преобразования адреса ячейки '{cell_addr}' при заполнении _display_data: {ve}")
                            
                logger.info(f"Заполнено {filled_cells_count} ячеек в модели для листа '{sheet_name}'.")
                
                self._sheet_name = sheet_name
                
            finally:
                self.endResetModel() # Завершаем сброс модели
                
            logger.info(f"Данные для листа '{sheet_name}' успешно загружены в модель.")

        except Exception as e:
            logger.error(f"Ошибка при загрузке данных для листа '{sheet_name}': {e}", exc_info=True)
            self._clear_data()
            self._sheet_name = sheet_name
            self.modelReset.emit() # Уведомить представление о сбросе модели

    def _clear_data(self):
        """Очищает внутренние данные модели."""
        self._display_data = []
        self._headers = []
        self._row_headers = []
        self._max_row = 0
        self._max_column = 0
        self._styles = {}
        self._sheet_name = None

    def rowCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = QModelIndex()) -> int:
        """Возвращает количество строк."""
        if parent.isValid():
            return 0
        # Возвращаем количество строк в данных или минимальное значение
        return len(self._display_data) if self._display_data else 0

    def columnCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = QModelIndex()) -> int:
        """Возвращает количество столбцов."""
        if parent.isValid():
            return 0
        # Возвращаем количество столбцов в данных или минимальное значение
        return len(self._display_data[0]) if self._display_data else 0

    def data(self, index: Union[QModelIndex, QPersistentModelIndex], role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        """Возвращает данные для указанной ячейки и роли."""
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if row >= len(self._display_data) or col >= len(self._display_data[0]):
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._display_data[row][col]
            # Возвращаем строковое представление значения
            return str(value) if value is not None else ""
        elif role == Qt.ItemDataRole.EditRole:
            # Для редактирования возвращаем "сырое" значение
            return self._display_data[row][col]
        # elif role == Qt.ItemDataRole.BackgroundRole:
        #     # Вернуть QBrush для фона ячейки на основе стиля
        #     # Пока не реализовано
        #     pass
        # elif role == Qt.ItemDataRole.ForegroundRole:
        #     # Вернуть QBrush для текста ячейки на основе стиля
        #     # Пока не реализовано
        #     pass
        # elif role == Qt.ItemDataRole.FontRole:
        #     # Вернуть QFont для ячейки на основе стиля
        #     # Пока не реализовано
        #     pass
        # elif role == Qt.ItemDataRole.TextAlignmentRole:
        #     # Вернуть флаги выравнивания на основе стиля
        #     # Пока не реализовано
        #     pass

        return None

    def setData(self, index: Union[QModelIndex, QPersistentModelIndex], value: Any, role: int = Qt.ItemDataRole.EditRole) -> bool:
        """
        Устанавливает данные для указанной ячейки.
        Вызывается при редактировании в QTableView.
        """
        logger.debug(f"TableModel.setData вызван для ({index.row()}, {index.column()}) со значением {value}")
        if not index.isValid() or role != Qt.ItemDataRole.EditRole:
            return False

        row = index.row()
        col = index.column()
        if row >= len(self._display_data) or col >= len(self._display_data[0]):
            return False

        cell_address = f"{_index_to_column_name(col)}{row + 1}"
        try:
            # Вызов AppController для обновления ячейки
            # Это синхронный вызов, который может быть медленным для больших операций.
            # В будущем можно рассмотреть использование QThread.
            if not self._sheet_name:
                 logger.error("Имя листа не установлено. Невозможно обновить ячейку.")
                 return False
                 
            success = self.app_controller.update_cell_value(self._sheet_name, cell_address, value)
            if success:
                # Обновляем локальное состояние модели
                self._display_data[row][col] = value
                # Уведомляем представление об изменении
                self.dataChanged.emit(index, index, [role])
                logger.info(f"Ячейка {cell_address} обновлена через AppController и модель.")
                return True
            else:
                logger.error(f"AppController не смог обновить ячейку {cell_address}.")
                return False
        except Exception as e:
            logger.error(f"Ошибка при обновлении ячейки {cell_address} через AppController: {e}", exc_info=True)
            return False

    def flags(self, index: Union[QModelIndex, QPersistentModelIndex]) -> Qt.ItemFlag:
        """Определяет флаги для ячейки (редактируемая, доступная и т.д.)."""
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags

        # Делаем ячейки редактируемыми
        return super().flags(index) | Qt.ItemFlag.ItemIsEditable

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

    # --- Дополнительные методы для работы с моделью ---
    def get_sheet_name(self) -> Optional[str]:
        """Возвращает имя текущего загруженного листа."""
        return self._sheet_name
        
    def get_cell_value(self, row: int, col: int) -> Any:
        """
        Получает значение ячейки по индексам.
        
        Args:
            row (int): Индекс строки (0-based).
            col (int): Индекс столбца (0-based).
            
        Returns:
            Any: Значение ячейки или None, если индексы вне диапазона.
        """
        if 0 <= row < len(self._display_data) and 0 <= col < len(self._display_data[0]):
            return self._display_data[row][col]
        return None
        
    def get_cell_address(self, row: int, col: int) -> str:
        """
        Преобразует индексы строки и столбца (0-based) в адрес ячейки Excel (e.g., 'A1').
        """
        col_name = _index_to_column_name(col)
        row_num = row + 1
        return f"{col_name}{row_num}"
    # ---------------------------------------------------
