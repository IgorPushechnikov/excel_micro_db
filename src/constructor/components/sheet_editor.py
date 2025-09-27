# src/constructor_new/components/sheet_editor.py
"""
Модуль для виджета редактора листа (новый GUI).
Использует QTableView с кастомной моделью и делегатом для отображения и редактирования данных Excel.
Интегрирован с AppController и DataManager.
"""

import logging
from typing import Optional, List, Dict, Any, Tuple

from PySide6.QtWidgets import QTableView, QAbstractItemDelegate, QStyledItemDelegate, QItemEditorFactory, QWidget, QVBoxLayout, QFrame, QAbstractItemView
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QPersistentModelIndex, Qt, Signal, QItemSelectionModel
from PySide6.QtGui import QColor, QBrush, QFont, QPainter

# Импортируем AppController
from src.core.app_controller import AppController
# Импортируем StyleManager и CellStyle
from .style_manager import StyleManager, CellStyle

# Получаем логгер
logger = logging.getLogger(__name__)


class SheetEditorModel(QAbstractTableModel):
    """
    Кастомная модель для QTableView, представляющая данные листа Excel.
    """

    def __init__(self, app_controller: AppController, sheet_name: str, parent=None):
        """
        Инициализирует модель.

        Args:
            app_controller (AppController): Экземпляр AppController.
            sheet_name (str): Имя листа, данные которого отображает модель.
            parent: Родительский объект.
        """
        super().__init__(parent)
        self.app_controller = app_controller
        self.sheet_name = sheet_name
        self._data: List[List[Any]] = []  # [[row1_col1, row1_col2, ...], [row2_col1, ...], ...]
        self._column_names: List[str] = [] # ["A", "B", "C", ...]
        # Инициализируем StyleManager
        self._style_manager = StyleManager(app_controller, sheet_name)
        self._load_data()

    def _load_data(self):
        """
        Загружает данные и стили для листа через AppController.
        """
        logger.info(f"Загрузка данных для листа '{self.sheet_name}' в модель...")
        try:
            # Используем get_sheet_editable_data, который возвращает структуру с 'column_names' и 'rows'
            sheet_data = self.app_controller.get_sheet_editable_data(self.sheet_name)
            if sheet_data:
                self._column_names = sheet_data.get("column_names", [])
                rows = sheet_data.get("rows", [])
                # rows - это список кортежей (tuple), конвертируем в список списков (list)
                self._data = [list(row) for row in rows]
                logger.info(f"Модель для листа '{self.sheet_name}' загружена: {len(self._data)} строк, {len(self._column_names)} столбцов.")
            else:
                logger.warning(f"Данные для листа '{self.sheet_name}' не загружены, используется пустая модель.")
                self._data = []
                self._column_names = ["A"] # Заглушка, хотя обычно должен быть хотя бы один столбец
        except Exception as e:
            logger.error(f"Ошибка при загрузке данных для листа '{self.sheet_name}': {e}", exc_info=True)
            self._data = []
            self._column_names = ["A"]

    def rowCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        """Возвращает количество строк."""
        if parent.isValid():
            return 0  # Упрощённая модель, без иерархии
        return len(self._data)

    def columnCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        """Возвращает количество столбцов."""
        if parent.isValid():
            return 0  # Упрощённая модель, без иерархии
        return len(self._column_names)

    def data(self, index: QModelIndex | QPersistentModelIndex, role: int = Qt.DisplayRole) -> Any:
        """Возвращает данные для указанной ячейки."""
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if row >= len(self._data) or col >= len(self._column_names):
            return None

        if role == Qt.DisplayRole or role == Qt.EditRole:
            return self._data[row][col] if col < len(self._data[row]) else ""

        # Роль для стиля (например, Qt.FontRole, Qt.BackgroundRole) будет обработана в делегате
        # или через itemData в делегате, но QStyledItemDelegate не всегда использует itemData напрямую для сложных стилей.
        # Лучше использовать делегат для отрисовки стилей.
        # Здесь мы можем вернуть данные стиля, но обычно это делает делегат, запрашивая их у модели.
        # Для простоты, пока не возвращаем стили здесь, а делегат будет запрашивать их у модели отдельно.
        # Однако, для QStyledItemDelegate, можно использовать setData и itemData, но это требует больше настройки.
        # Пока оставим стили на делегате, который будет запрашивать их через отдельный метод.
        # Создадим такой метод ниже.

        return None

    def setData(self, index: QModelIndex | QPersistentModelIndex, value: Any, role: int = Qt.EditRole) -> bool:
        """Устанавливает данные для указанной ячейки."""
        if not index.isValid() or role != Qt.EditRole:
            return False

        row = index.row()
        col = index.column()

        if row >= len(self._data) or col >= len(self._column_names):
            return False

        # Обновляем значение в модели
        # Убедимся, что строка достаточно длинна
        while col >= len(self._data[row]):
            self._data[row].append("")

        old_value = self._data[row][col]
        self._data[row][col] = value

        # Уведомляем представление об изменении
        self.dataChanged.emit(index, index, [role])

        # Отправляем изменение в AppController
        # Для этого нужно преобразовать row/col в имя столбца Excel (A, B, ...)
        column_name = self._column_names[col] # Предполагаем, что _column_names содержит A, B, C...
        success = self.app_controller.update_sheet_cell_in_project(self.sheet_name, row, column_name, value)
        if not success:
             logger.error(f"Не удалось обновить ячейку {column_name}{row+1} в проекте.")
             # Возвращаем старое значение?
             # self._data[row][col] = old_value
             # self.dataChanged.emit(index, index, [role])
             return False

        logger.debug(f"Ячейка ({row}, {col}) обновлена с '{old_value}' на '{value}'.")
        return True

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        """Возвращает данные для заголовков строк/столбцов."""
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                if section < len(self._column_names):
                    return self._column_names[section]
                else:
                    return "" # Или None?
            elif orientation == Qt.Vertical:
                return str(section + 1) # Нумерация строк с 1
        return None

    def flags(self, index: QModelIndex | QPersistentModelIndex) -> Qt.ItemFlags:
        """Возвращает флаги для ячейки (редактируемость, выбор и т.д.)."""
        if not index.isValid():
            return Qt.NoItemFlags

        return super().flags(index) | Qt.ItemIsEditable # Делает ячейки редактируемыми

    # --- Метод для получения стиля ячейки (для делегата) ---
    def get_style_for_cell(self, row: int, col: int) -> Optional[Dict[str, Any]]:
        """
        Возвращает словарь стиля для ячейки (row, col).
        Формат словаря: {"font": QFont, "bg_color": QColor, "text_color": QColor, ...}
        """
        cell_style = self._style_manager.get_style(row, col)
        if cell_style is None:
            return None

        # Конвертируем CellStyle в словарь для совместимости с делегатом
        style_dict = {}
        if cell_style.bg_color:
            style_dict["bg_color"] = cell_style.bg_color
        if cell_style.text_color:
            style_dict["text_color"] = cell_style.text_color
        if cell_style.font:
            style_dict["font"] = cell_style.font
        return style_dict if style_dict else None


class SheetEditorDelegate(QStyledItemDelegate):
    """
    Кастомный делегат для QTableView, отвечающий за отрисовку ячеек с учетом стилей.
    """
    def __init__(self, model: SheetEditorModel, parent=None):
        """
        Инициализирует делегат.

        Args:
            model (SheetEditorModel): Модель, с которой связан делегат.
            parent: Родительский объект.
        """
        super().__init__(parent)
        self.model = model

    def paint(self, painter: QPainter, option, index: QModelIndex | QPersistentModelIndex):
        """
        Отрисовывает ячейку с учетом стиля.
        """
        # Получаем стиль для текущей ячейки из модели
        row = index.row()
        col = index.column()
        style_dict = self.model.get_style_for_cell(row, col)

        # Модифицируем option перед стандартной отрисовкой
        if style_dict:
            # Установка фонового цвета через палитру
            bg_color = style_dict.get("bg_color")
            if bg_color:
                option.palette.setColor(option.palette.Base, bg_color)

            # Установка цвета текста через палитру
            text_color = style_dict.get("text_color")
            if text_color:
                option.palette.setColor(option.palette.Text, text_color)

            # Установка шрифта
            font = style_dict.get("font")
            if font:
                option.font = font

        # Вызываем стандартную отрисовку с модифицированным option
        super().paint(painter, option, index)


class SheetEditor(QFrame):
    """
    Виджет редактора листа (новый GUI).
    """
    # Сигнал, эмитируемый при изменении выделения ячейки
    cellSelectionChanged = Signal(str, str) # (cell_address, cell_value)

    def __init__(self, app_controller: AppController, sheet_name: str, parent=None):
        """
        Инициализирует редактор листа.

        Args:
            app_controller (AppController): Экземпляр основного контроллера приложения.
            sheet_name (str): Имя листа для редактирования.
            parent: Родительский объект.
        """
        super().__init__(parent)
        self.app_controller = app_controller
        self.sheet_name = sheet_name
        self._setup_ui()

    def set_cell_style(self, row: int, col: int, style: 'CellStyle') -> bool:
        """
        Устанавливает стиль для ячейки (row, col).

        Args:
            row (int): Номер строки (0-базированный).
            col (int): Номер столбца (0-базированный).
            style (CellStyle): Стиль для установки.

        Returns:
            bool: True, если стиль успешно установлен, иначе False.
        """
        # Делегируем установку стиля модели
        success = self.model._style_manager.set_style(row, col, style)
        if success:
            # Обновляем представление для отображения нового стиля
            index = self.model.index(row, col)
            self.model.dataChanged.emit(index, index, [Qt.DisplayRole])
        return success

    def _setup_ui(self):
        """
        Настраивает пользовательский интерфейс.
        """
        layout = QVBoxLayout(self)
        self.table_view = QTableView()

        # Создаём и устанавливаем модель
        self.model = SheetEditorModel(self.app_controller, self.sheet_name, self)
        self.table_view.setModel(self.model)

        # Создаём и устанавливаем делегат
        self.delegate = SheetEditorDelegate(self.model, self.table_view)
        self.table_view.setItemDelegate(self.delegate)

        # Настройки представления
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectItems) # Выделение ячеек
        self.table_view.setSelectionMode(QAbstractItemView.SingleSelection) # Одиночное выделение
        # self.table_view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed) # Триггеры редактирования

        # Подключение сигнала изменения выделения для обновления строки формул
        self.table_view.selectionModel().selectionChanged.connect(self._on_selection_changed)

        layout.addWidget(self.table_view)

    def _on_selection_changed(self, selected, deselected):
        """
        Обработчик изменения выделения ячейки.
        """
        # Находим индекс новой выделенной ячейки
        current_index = self.table_view.currentIndex()
        if current_index.isValid():
            # Получаем адрес ячейки (например, A1)
            # Для этого нужно преобразовать индекс в формат Excel (A1, B2 и т.д.)
            row = current_index.row()
            col = current_index.column()
            # Пример функции преобразования: _col_idx_to_name
            col_name = self._col_idx_to_name(col)
            cell_address = f"{col_name}{row + 1}"
            # Получаем значение ячейки
            cell_value = current_index.data(Qt.DisplayRole)
            if cell_value is None:
                cell_value = ""
            # Сообщаем главное окно (или GUIController) об изменении
            # Это можно сделать через сигнал или напрямую, если MainWindow знает о SheetEditor
            # Пока просто логируем.
            logger.debug(f"Выделена ячейка: {cell_address}, значение: '{cell_value}'")
            # Эмитим сигнал
            self.cellSelectionChanged.emit(cell_address, cell_value)

    def _col_idx_to_name(self, idx: int) -> str:
        """
        Преобразует индекс столбца (0-базированный) в имя столбца Excel (A, B, ..., Z, AA, AB, ...).
        """
        name = ""
        while idx >= 0:
            name = chr(idx % 26 + ord('A')) + name
            idx = idx // 26 - 1
        return name
