# backend/constructor/widgets/sheet_editor.py
"""
Виджет-редактор для отображения и редактирования содержимого листа Excel.
"""

import sys
import string  # Для генерации имен столбцов Excel
import json
import re

# ИСПРАВЛЕНО: Добавлен QPersistentModelIndex в импорт
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, Slot, Signal, QPersistentModelIndex
from PySide6.QtGui import QBrush, QColor, QAction, QFont # Добавлен QFont для стилей шрифта
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTableView, QLabel, QMessageBox,
    QAbstractItemView, QHeaderView, QApplication, QMenu, QInputDialog, QHBoxLayout, QLineEdit
)

# Импорты для типизации
from typing import Optional, Dict, Any, List, NamedTuple, Union
from pathlib import Path

import sqlite3
import logging

# Импорт для аннотаций типов, избегая циклических импортов
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    # ИСПРАВЛЕНО: Импорт AppController теперь из core внутри backend
    from core.app_controller import AppController

# ИСПРАВЛЕНО: Импорт logger теперь из utils внутри backend
from utils.logger import get_logger

# --- НОВОЕ: Импорт модуля форматирования ---
from .cell_formatting import format_cell_value
# ---------------------------------------------

logger = get_logger(__name__)


# === НОВОЕ: Структура для хранения информации об одном редактировании ===
class EditAction(NamedTuple):
    """Представляет одно действие редактирования для Undo/Redo."""
    row: int
    col: int
    old_value: Any
    new_value: Any


# =====================================================================
# === ИЗМЕНЕНО: SheetDataModel с исправленными сигнатурами и логикой ===
class SheetDataModel(QAbstractTableModel):
    """
    Модель данных для отображения и редактирования содержимого листа в QTableView.
    Отображает данные как в Excel: первая строка данных - это данные,
    заголовки столбцов - стандартные имена Excel (A, B, C...).
    """

    # === НОВОЕ: Сигнал, испускаемый ДО изменения данных ===
    # QModelIndex, старое значение, новое значение
    cellDataAboutToChange = Signal(QModelIndex, object, object)
    # ===================================================

    # === СУЩЕСТВУЮЩИЙ: Сигнал для внутреннего использования ===
    dataChangedExternally = Signal(QModelIndex, QModelIndex)
    # =========================================================

    def __init__(self, editable_data: Dict[str, Any], parent=None):
        super().__init__(parent)
        self._editable_data = editable_data
        # Сохраняем оригинальные имена столбцов (из первой строки Excel)
        self._original_column_names = self._editable_data.get("column_names", [])
        # Данные ячеек
        raw_rows = self._editable_data.get("rows", [])
        self._rows: List[List[Any]] = [list(row_tuple) for row_tuple in raw_rows]

        # === НОВОЕ: Хранение стилей для ячеек ===
        self._cell_styles: Dict[tuple, Dict[str, Any]] = {}
        # ======================================

        # Генерируем стандартные имена столбцов Excel
        self._generated_column_headers = self._generate_excel_column_names(
            len(self._rows[0]) if self._rows else 0
        )

    def set_cell_styles(self, styles_data: List[Dict[str, Any]]):
        """Устанавливает стили для ячеек на основе данных из storage.styles."""
        self._cell_styles.clear()
        logger.debug(f"SheetDataModel.set_cell_styles: Получено {len(styles_data)} записей стилей.")
        for style_info in styles_data:
            range_addr = style_info.get("range_address", "")
            style_attrs_json = style_info.get("style_attributes", "{}") # По умолчанию пустой JSON
            if range_addr:
                # --- Парсим JSON ---
                try:
                    style_attrs = json.loads(style_attrs_json) if isinstance(style_attrs_json, str) else style_attrs_json
                    if not isinstance(style_attrs, dict):
                         logger.warning(f"SheetDataModel.set_cell_styles: 'style_attributes' для '{range_addr}' не является словарем или корректным JSON-объектом. Пропущено.")
                         continue
                except json.JSONDecodeError as e:
                    logger.warning(f"SheetDataModel.set_cell_styles: Ошибка парсинга JSON стиля для '{range_addr}': {e}")
                    continue # Пропускаем некорректный стиль

                # --- Парсер адреса диапазона Excel ---
                # Регулярное выражение для адреса одной ячейки или диапазона
                # Примеры: 'A1', 'Z10', 'AA1', 'A1:B2', 'ZZ100:AAA200'
                # Группы: 1-столбец_начала, 2-строка_начала, 3-столбец_конца, 4-строка_конца (опционально)
                range_pattern = re.compile(r'^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$')
                match = range_pattern.match(range_addr)
                if not match:
                    logger.warning(f"SheetDataModel.set_cell_styles: Не удалось распознать формат адреса '{range_addr}'. Пропущено.")
                    continue

                start_col_letter, start_row_str, end_col_letter, end_row_str = match.groups()

                # Если это одиночная ячейка, end_... будет None
                if end_col_letter is None or end_row_str is None:
                    end_col_letter = start_col_letter
                    end_row_str = start_row_str

                try:
                    start_col_index = self._column_letter_to_index(start_col_letter)
                    start_row_index = int(start_row_str) - 1  # Excel 1-based -> Python 0-based
                    end_col_index = self._column_letter_to_index(end_col_letter)
                    end_row_index = int(end_row_str) - 1      # Excel 1-based -> Python 0-based

                    if start_col_index < 0 or start_row_index < 0 or end_col_index < 0 or end_row_index < 0:
                        logger.warning(f"SheetDataModel.set_cell_styles: Некорректный индекс после парсинга {range_addr}: [{start_row_index}, {start_col_index}] - [{end_row_index}, {end_col_index}]")
                        continue

                    # Применяем стиль ко всем ячейкам в диапазоне
                    for r in range(start_row_index, end_row_index + 1):
                        for c in range(start_col_index, end_col_index + 1):
                            # Логируем загруженные стили для отладки (опционально, можно убрать)
                            # logger.debug(f"SheetDataModel.set_cell_styles: Установлен стиль для [{r}, {c}]: {list(style_attrs.keys())}")
                            self._cell_styles[(r, c)] = style_attrs

                    logger.debug(f"SheetDataModel.set_cell_styles: Применен стиль к диапазону {range_addr} ([{start_row_index}, {start_col_index}] - [{end_row_index}, {end_col_index}]).")

                except (ValueError, TypeError) as e:
                    logger.warning(f"SheetDataModel.set_cell_styles: Ошибка преобразования индексов для '{range_addr}': {e}")
                    continue

        # Сообщаем представлению, что данные могли измениться (для перерисовки стилей)
        if self._rows:
            top_left = self.index(0, 0)
            bottom_right = self.index(self.rowCount() - 1, self.columnCount() - 1)
            if top_left.isValid() and bottom_right.isValid():
                self.dataChanged.emit(top_left, bottom_right)
                logger.debug("SheetDataModel.set_cell_styles: Сигнал dataChanged отправлен для обновления стилей.")

    def _column_letter_to_index(self, letter: str) -> int:
        """Преобразует букву столбца Excel (например, 'A', 'Z', 'AA') в 0-базовый индекс."""
        result = 0
        for char in letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1  # 0-based index

    def _column_index_to_letter(self, index: int) -> str:
        """
        Преобразует 0-базовый индекс столбца в букву Excel (например, 0 -> 'A', 25 -> 'Z', 26 -> 'AA').
        """
        if index < 0:
            return ""
        result = ""
        while index >= 0:
            result = chr(index % 26 + ord('A')) + result
            index = index // 26 - 1
        return result

    def _generate_excel_column_names(self, count: int) -> List[str]:
        """Генерирует список имен столбцов Excel (A, B, ..., Z, AA, AB, ...)."""
        names = []
        for i in range(count):
            name = ""
            temp = i
            while temp >= 0:
                name = string.ascii_uppercase[temp % 26] + name
                temp = temp // 26 - 1
                if temp < 0:
                    break
            names.append(name if name else "A")  # fallback для count=0
        return names

    # ИСПРАВЛЕНО СНОВА: Сигнатура метода rowCount соответствует базовому классу строго
    # Pylance требует Union[QModelIndex, QPersistentModelIndex]
    def rowCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = ...) -> int: # type: ignore
        _ = parent # Чтобы избежать предупреждения о неиспользованном параметре
        return len(self._rows)

    # ИСПРАВЛЕНО СНОВА: Сигнатура метода columnCount соответствует базовому классу строго
    # Pylance требует Union[QModelIndex, QPersistentModelIndex]
    def columnCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = ...) -> int: # type: ignore
        _ = parent # Чтобы избежать предупреждения о неиспользованном параметре
        return len(self._rows[0]) if self._rows else 0

    # ИСПРАВЛЕНО СНОВА: Сигнатура метода data соответствует базовому классу строго
    # Pylance требует Union[QModelIndex, QPersistentModelIndex] для index
    def data(self, index: Union[QModelIndex, QPersistentModelIndex], role: int = ...) -> Any: # type: ignore
        # Проверка валидности индекса
        if not index.isValid():
            return None

        # НЕ НУЖНО преобразовывать QPersistentModelIndex, так как QModelIndex может его представлять
        # if isinstance(index, QPersistentModelIndex):
        #     index = QModelIndex(index)

        row = index.row()
        col = index.column()

        if role == Qt.ItemDataRole.DisplayRole:
            if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
                value = self._rows[row][col]
                # --- ИЗМЕНЕНО: Применение форматирования ---
                # Получаем стиль для ячейки
                style = self._cell_styles.get((row, col))
                number_format_code = style.get("number_format") if style else None
                # Используем новую функцию для форматирования
                formatted_value = format_cell_value(value, number_format_code)
                # logger.debug(f"Форматирование: [{row},{col}] {value} -> {formatted_value} по формату '{number_format_code}'")
                return formatted_value if formatted_value is not None else ""
                # --- КОНЕЦ ИЗМЕНЕНИЯ ---
        elif role == Qt.ItemDataRole.ToolTipRole:
            if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
                value = self._rows[row][col]
                # Для ToolTip используем оригинальное значение, возможно, отформатированное проще
                orig_name = self._original_column_names[col] if col < len(self._original_column_names) else f"Col_{col}"
                return f"Столбец: {orig_name}\nЗначение: {repr(value)}"
        # === НОВОЕ: Обработка ролей для стилей ===
        elif role == Qt.ItemDataRole.BackgroundRole:
            style = self._cell_styles.get((row, col))
            if style:
                # logger.debug(f"Запрос BackgroundRole для [{row},{col}]: {style}") # Для отладки
                # OpenPyXL обычно хранит цвет фона в fill_fg_color
                fill_fg_color = style.get("fill_fg_color")
                # Также проверим fill_bg_color, если fg_color не задан или пуст
                if not fill_fg_color or fill_fg_color == "00000000": # Прозрачный
                     fill_fg_color = style.get("fill_bg_color")
                if fill_fg_color and isinstance(fill_fg_color, str) and fill_fg_color.startswith('#'):
                    try:
                        color = QColor(fill_fg_color)
                        if color.isValid():
                            # logger.debug(f"BackgroundRole: Цвет {fill_fg_color} применен к [{row},{col}]")
                            return QBrush(color)
                        else:
                            logger.debug(f"BackgroundRole: Цвет {fill_fg_color} невалиден для [{row},{col}]")
                    except Exception as e:
                        logger.warning(f"Ошибка преобразования цвета фона '{fill_fg_color}' для [{row},{col}]: {e}")
        elif role == Qt.ItemDataRole.ForegroundRole:
            style = self._cell_styles.get((row, col))
            if style:
                # logger.debug(f"Запрос ForegroundRole для [{row},{col}]: {style}") # Для отладки
                # OpenPyXL обычно хранит цвет шрифта в font_color
                font_color = style.get("font_color")
                if font_color and isinstance(font_color, str) and font_color.startswith('#'):
                    try:
                        color = QColor(font_color)
                        if color.isValid():
                            # logger.debug(f"ForegroundRole: Цвет {font_color} применен к [{row},{col}]")
                            return QBrush(color)
                        else:
                           logger.debug(f"ForegroundRole: Цвет {font_color} невалиден для [{row},{col}]")
                    except Exception as e:
                        logger.warning(f"Ошибка преобразования цвета шрифта '{font_color}' для [{row},{col}]: {e}")
        elif role == Qt.ItemDataRole.FontRole:
             # Можно добавить обработку стилей шрифта (жирность, курсив и т.д.) здесь
             style = self._cell_styles.get((row, col))
             if style:
                 font = QFont()
                 # Жирный шрифт
                 font_b = style.get("font_b")
                 if font_b is True or (isinstance(font_b, str) and font_b.lower() in ('1', 'true')):
                     font.setBold(True)
                 # Курсив
                 font_i = style.get("font_i")
                 if font_i is True or (isinstance(font_i, str) and font_i.lower() in ('1', 'true')):
                     font.setItalic(True)
                 # Подчеркнутый (u может быть 'single', 'double' и т.д.)
                 font_u = style.get("font_u")
                 if font_u and font_u != "none":
                     font.setUnderline(True)
                 # Зачеркнутый
                 font_strike = style.get("font_strike")
                 if font_strike is True or (isinstance(font_strike, str) and font_strike.lower() in ('1', 'true')):
                     font.setStrikeOut(True)
                 return font
        # ======================================

        return None

    # ИСПРАВЛЕНО СНОВА: Сигнатура метода headerData соответствует базовому классу строго
    # Pylance требует int для role
    def headerData(self, section: int, orientation: Qt.Orientation, role: int = ...) -> Any: # type: ignore
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                if 0 <= section < len(self._generated_column_headers):
                    return self._generated_column_headers[section]
                else:
                    return f"Col_{section}"  # fallback
            elif orientation == Qt.Orientation.Vertical:
                # Номера строк (1-based), как в Excel
                return str(section + 1)
        return None

    # ИСПРАВЛЕНО СНОВА: Сигнатура метода flags соответствует базовому классу строго
    # Pylance требует Union[QModelIndex, QPersistentModelIndex] для index
    def flags(self, index: Union[QModelIndex, QPersistentModelIndex]) -> Qt.ItemFlag: # type: ignore
        # НЕ НУЖНО преобразовывать QPersistentModelIndex
        # if isinstance(index, QPersistentModelIndex):
        #     index = QModelIndex(index)

        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled

    # ИСПРАВЛЕНО СНОВА: Сигнатура метода setData соответствует базовому классу строго
    # Pylance требует Union[QModelIndex, QPersistentModelIndex] для index
    def setData(self, index: Union[QModelIndex, QPersistentModelIndex], value: Any, role: int = ...) -> bool: # type: ignore
        """
        Устанавливает данные в модель. Вызывается, когда пользователь редактирует ячейку.
        Испускает cellDataAboutToChange до изменения и dataChanged после.
        """
        # НЕ НУЖНО преобразовывать QPersistentModelIndex
        # if isinstance(index, QPersistentModelIndex):
        #     index = QModelIndex(index)

        if index.isValid() and role == Qt.ItemDataRole.EditRole:
            row = index.row()
            col = index.column()
            if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
                old_value = self._rows[row][col]
                new_value_str = str(value) if value is not None else ""
                logger.debug(
                    f"SheetDataModel: Испускание cellDataAboutToChange для [{row},{col}]: '{old_value}' -> '{new_value_str}'")
                self.cellDataAboutToChange.emit(index, old_value, new_value_str)
                self._rows[row][col] = new_value_str
                # role передается как список в dataChanged
                self.dataChanged.emit(index, index, [role])
                logger.debug(
                    f"SheetDataModel: Данные ячейки [{row}, {col}] изменены с '{old_value}' на '{new_value_str}'.")
                return True
        return False

    # =========================================================
    def setDataInternal(self, row: int, col: int, value):
        """Внутренне обновляет данные модели без вызова setData."""
        if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
            self._rows[row][col] = value
            index = self.index(row, col)
            if index.isValid():
                self.dataChangedExternally.emit(index, index)
            logger.debug(f"Модель (внутр.): Данные ячейки [{row}, {col}] обновлены до '{value}'.")


# === SheetEditor с исправлениями и улучшениями ===
class SheetEditor(QWidget):
    """
    Виджет для редактирования/просмотра содержимого одного листа.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project_db_path: Optional[str] = None
        self.sheet_name: Optional[str] = None
        self._model: Optional[SheetDataModel] = None
        self.app_controller: Optional['AppController'] = None

        # === Стеки для Undo/Redo ===
        self._undo_stack: List[EditAction] = []
        self._redo_stack: List[EditAction] = []
        self._max_undo_steps = 50
        # ==========================

        # === НОВОЕ: Строка редактирования и индикация ячейки ===
        self._formula_bar: Optional[QLineEdit] = None
        self._current_editing_index: Optional[QModelIndex] = None
        self._cell_address_label: Optional[QLabel] = None # Для отображения адреса активной ячейки
        # ===================================

        # === УДАЛЕНО: Подключение к несуществующему сигналу selectionModelChanged ===
        # self.table_view.selectionModelChanged.connect(self._on_selection_model_changed)
        # ======================================================================

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.label_sheet_name = QLabel("Лист: <Не выбран>")
        self.label_sheet_name.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(self.label_sheet_name)

        # === НОВОЕ: Добавление строки редактирования и индикации ячейки ===
        formula_layout = QHBoxLayout()
        self._cell_address_label = QLabel("Ячейка: ")
        self._cell_address_label.setStyleSheet("font-weight: normal; padding: 2px;")
        self._cell_address_label.setMinimumWidth(80) # Минимальная ширина для адреса

        formula_label = QLabel("Формула/Значение:")
        self._formula_bar = QLineEdit()
        # Проверка на None не обязательна здесь, так как мы только что создали объект
        if self._formula_bar:
             self._formula_bar.setPlaceholderText("Выберите ячейку или введите значение/формулу")
             self._formula_bar.returnPressed.connect(self._on_formula_bar_return_pressed)
             # Подключаем сигнал изменения текста для обновления модели при потере фокуса или Enter
             # self._formula_bar.editingFinished.connect(self._on_formula_bar_editing_finished) # Альтернатива
        formula_layout.addWidget(self._cell_address_label)
        formula_layout.addWidget(formula_label)
        formula_layout.addWidget(self._formula_bar)
        layout.addLayout(formula_layout)
        # =============================================

        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        # ИЗМЕНЕНО: Позволяем выбирать только одну ячейку для упрощения логики
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection) # Только одна ячейка

        # === ИСПРАВЛЕНО: Подключение сигнала clicked ===
        self.table_view.clicked.connect(self._on_cell_clicked)
        # === УДАЛЕНО: Подключение selectionChanged из _setup_ui ===
        # self.table_view.selectionModel().selectionChanged.connect(self._on_selection_changed)
        # Подключение selectionChanged теперь происходит в _connect_selection_model_signals
        # =============================================

        # Настройка контекстного меню
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self._on_context_menu)

        # === ИЗМЕНЕНО: Настройка заголовков ===
        # Горизонтальный заголовок (столбцы)
        horizontal_header = self.table_view.horizontalHeader()
        horizontal_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive) # Позволяем пользователю менять ширину
        # Вертикальный заголовок (строки)
        vertical_header = self.table_view.verticalHeader()
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.Fixed) # Фиксированная высота строк
        vertical_header.setDefaultSectionSize(20) # Высота строки по умолчанию

        layout.addWidget(self.table_view)

        # Создание действий для Undo/Redo
        self.action_undo = QAction("Отменить", self)
        self.action_undo.setShortcut("Ctrl+Z")
        self.action_undo.triggered.connect(self.undo)
        self.action_undo.setEnabled(False)

        self.action_redo = QAction("Повторить", self)
        self.action_redo.setShortcut("Ctrl+Y")
        self.action_redo.triggered.connect(self.redo)
        self.action_redo.setEnabled(False)

    @Slot(QModelIndex)
    def _on_cell_clicked(self, index: QModelIndex):
        """Обработчик клика по ячейке в таблице."""
        # logger.debug(f"SheetEditor._on_cell_clicked: Клик по индексу {index.row()}, {index.column()}")
        self._update_formula_bar(index)

    # === УДАЛЕНО: Слот _on_selection_model_changed ===
    # @Slot() # Для selectionModelChanged
    # def _on_selection_model_changed(self):
    #     """Слот, вызываемый при изменении модели выделения таблицы (обычно после setModel)."""
    #     logger.debug("SheetEditor._on_selection_model_changed: Модель выделения изменилась.")
    #     new_selection_model = self.table_view.selectionModel()
    #     if new_selection_model:
    #         # Подключаемся к сигналу selectionChanged новой модели выделения
    #         new_selection_model.selectionChanged.connect(self._on_selection_changed)
    #         logger.debug("SheetEditor._on_selection_model_changed: Подключен к selectionChanged новой модели.")
    #     else:
    #         logger.warning("SheetEditor._on_selection_model_changed: Новая модель выделения - None.")
    # ==============================================================

    @Slot() # Для selectionChanged
    def _on_selection_changed(self):
        """Обработчик изменения выделения в таблице."""
        # logger.debug("SheetEditor._on_selection_changed: Выделение изменилось")
        # === ИЗМЕНЕНО: Проверка на None и проверка наличия selectedIndexes ===
        selection_model = self.table_view.selectionModel()
        if selection_model:
            selected_indexes = selection_model.selectedIndexes()
            if selected_indexes:
                # Берем первую выбранную ячейку
                index = selected_indexes[0]
                self._update_formula_bar(index)
            else:
                # Если ничего не выбрано, очищаем строку редактирования
                self._clear_formula_bar()
        # ==============================================================

    def _update_formula_bar(self, index: QModelIndex):
        """Обновляет строку редактирования и метку адреса на основе выбранного индекса."""
        if not index.isValid() or not self._model:
            self._clear_formula_bar()
            return

        self._current_editing_index = QModelIndex(index) # Сохраняем копию индекса

        # Обновляем метку с адресом ячейки
        if self._cell_address_label:
            row = index.row()
            col = index.column()
            # Генерируем имя столбца Excel
            col_name = self._model._generated_column_headers[col] if 0 <= col < len(self._model._generated_column_headers) else f"Col_{col}"
            # Номер строки в Excel 1-based
            row_name = str(row + 1)
            cell_address = f"{col_name}{row_name}"
            self._cell_address_label.setText(f"Ячейка: {cell_address}")

        # Получаем значение ячейки из модели для отображения в строке редактирования
        display_value = self._model.data(index, Qt.ItemDataRole.DisplayRole)
        # Проверяем, что _formula_bar существует
        if self._formula_bar:
            # Убеждаемся, что display_value - строка
            text_to_set = ""
            if display_value is not None:
                if isinstance(display_value, str):
                    text_to_set = display_value
                else:
                    text_to_set = str(display_value)
            self._formula_bar.setText(text_to_set)
            # logger.debug(f"SheetEditor._update_formula_bar: Установлен текст '{text_to_set}' для ячейки {index.row()}, {index.column()}")

    def _clear_formula_bar(self):
        """Очищает строку редактирования и метку адреса."""
        self._current_editing_index = None
        if self._cell_address_label:
            self._cell_address_label.setText("Ячейка: ")
        if self._formula_bar:
            self._formula_bar.setText("")
            self._formula_bar.setPlaceholderText("Выберите ячейку или введите значение/формулу")

    @Slot()
    def _on_formula_bar_return_pressed(self):
        """Обработчик нажатия Enter в строке редактирования."""
        self._apply_formula_bar_value()

    # @Slot() # Альтернатива для editingFinished
    # def _on_formula_bar_editing_finished(self):
    #     """Обработчик завершения редактирования строки редактирования."""
    #     self._apply_formula_bar_value()

    def _apply_formula_bar_value(self):
        """Применяет значение из строки редактирования к выбранной ячейке."""
        # Проверяем, что _formula_bar существует
        if not self._formula_bar:
            logger.warning("Строка редактирования не инициализирована.")
            return

        if not self._current_editing_index or not self._current_editing_index.isValid() or not self._model:
            logger.debug("Строка редактирования: Нет выбранной ячейки для обновления.")
            return

        new_text = self._formula_bar.text()
        # Устанавливаем данные в модель
        success = self._model.setData(self._current_editing_index, new_text, Qt.ItemDataRole.EditRole)
        if success:
            logger.debug(f"Строка редактирования: Значение '{new_text}' установлено в ячейку.")
            # Явно выбираем ячейку в таблице после обновления, чтобы сохранить выделение
            self.table_view.setCurrentIndex(self._current_editing_index)
        else:
            logger.warning(f"Строка редактирования: Не удалось установить значение '{new_text}' в ячейку.")

    @Slot(object) # object для position
    def _on_context_menu(self, position):
        """Создает и показывает контекстное меню для таблицы."""
        context_menu = QMenu(self)
        context_menu.addAction(self.action_undo)
        context_menu.addAction(self.action_redo)
        # context_menu.addAction(...) для других действий
        context_menu.exec(self.table_view.viewport().mapToGlobal(position))

    def set_app_controller(self, controller: Optional['AppController']):
        self.app_controller = controller
        logger.debug("SheetEditor: AppController установлен/обновлён.")

    def load_sheet(self, project_db_path: str, sheet_name: str):
        logger.info(f"Загрузка листа '{sheet_name}' из БД: {project_db_path}")
        self.project_db_path = project_db_path
        self.sheet_name = sheet_name
        self.label_sheet_name.setText(f"Лист: {sheet_name}")

        # Очистка стеков Undo/Redo при загрузке нового листа
        self._clear_undo_redo_stacks()

        # === НОВОЕ: Очистка строки редактирования ===
        self._clear_formula_bar() # Используем общий метод очистки
        # =========================================

        if not self.app_controller:
            logger.error("SheetEditor: AppController не установлен для загрузки листа.")
            QMessageBox.critical(self, "Ошибка", "Контроллер приложения не доступен для загрузки данных.")
            return

        try:
            editable_data = self.app_controller.get_sheet_editable_data(sheet_name)
            if editable_data is not None and 'column_names' in editable_data:
                self._model = SheetDataModel(editable_data)
                # === НОВОЕ: Подключение к новому сигналу модели ===
                self._model.cellDataAboutToChange.connect(self._on_cell_data_about_to_change)
                # =================================================
                # --- ИСПРАВЛЕНО: Подключение сигнала dataChanged ---
                self._model.dataChanged.connect(self._on_model_data_changed)
                # ==================================================
                self._model.dataChangedExternally.connect(self.table_view.dataChanged)
                # --- ЭТОТ ВЫЗОВ ВАЖЕН ---
                self.table_view.setModel(self._model)
                # ------------------------
                # === НОВОЕ: Подключение сигналов модели выделения ПОСЛЕ setModel ===
                self._connect_selection_model_signals()
                # =================================================================
                logger.info(f"Лист '{sheet_name}' успешно загружен в редактор. "
                            f"Строк: {len(editable_data.get('rows', []))}, "
                            f"Столбцов: {len(editable_data.get('column_names', []))}")

                # === НОВОЕ: Загрузка стилей и применение к модели ===
                try:
                    conn = sqlite3.connect(self.project_db_path)
                    cursor = conn.cursor()
                    cursor.execute("SELECT id FROM sheets WHERE name = ?", (sheet_name,))
                    sheet_row = cursor.fetchone()
                    if sheet_row:
                        sheet_id = sheet_row[0]
                        # Импортируем функцию загрузки стилей
                        # ИСПРАВЛЕНО: Путь к load_sheet_styles теперь storage.styles
                        from storage.styles import load_sheet_styles
                        styles_data = load_sheet_styles(conn, sheet_id)
                        logger.info(f"SheetEditor.load_sheet: Загружено {len(styles_data)} стилей для листа '{sheet_name}' из БД.")
                        # Передаем стили в модель
                        self._model.set_cell_styles(styles_data)
                    else:
                        logger.warning(f"SheetEditor.load_sheet: Лист '{sheet_name}' не найден в БД при попытке загрузить стили.")
                    conn.close()
                except Exception as e:
                    logger.error(f"SheetEditor.load_sheet: Ошибка при загрузке стилей для листа '{sheet_name}': {e}", exc_info=True)
                # =================================================

            else:
                self.table_view.setModel(None) # Это должно вызвать selectionModelChanged
                self._model = None
                # === НОВОЕ: Отключаем сигналы модели выделения при очистке ===
                self._disconnect_selection_model_signals()
                # =========================================================
                logger.warning(f"SheetEditor.load_sheet: Редактируемые данные для листа '{sheet_name}' не найдены или пусты.")
        except Exception as e:
            logger.error(f"SheetEditor.load_sheet: Ошибка при загрузке листа '{sheet_name}': {e}", exc_info=True)
            QMessageBox.critical(
                self,
                "Ошибка загрузки",
                f"Не удалось загрузить содержимое листа '{sheet_name}':\n{e}"
            )
            self.table_view.setModel(None) # Это тоже должно вызвать selectionModelChanged
            self._model = None
            # === НОВОЕ: Отключаем сигналы модели выделения при ошибке ===
            self._disconnect_selection_model_signals()
            # =========================================================

    # === НОВОЕ: Метод для подключения сигналов модели выделения ===
    def _connect_selection_model_signals(self):
        """Подключает сигналы модели выделения таблицы."""
        selection_model = self.table_view.selectionModel()
        if selection_model:
            selection_model.selectionChanged.connect(self._on_selection_changed)
            logger.debug("SheetEditor: Подключен к сигналу selectionChanged модели выделения.")
        else:
            logger.warning("SheetEditor: Модель выделения отсутствует при попытке подключения сигналов.")

    # === НОВОЕ: Метод для отключения сигналов модели выделения ===
    def _disconnect_selection_model_signals(self):
        """Отключает сигналы модели выделения таблицы."""
        selection_model = self.table_view.selectionModel()
        if selection_model:
            # Отключаем сигнал, если он был подключен
            try:
                selection_model.selectionChanged.disconnect(self._on_selection_changed)
                logger.debug("SheetEditor: Отключен от сигнала selectionChanged модели выделения.")
            except RuntimeError:
                # Сигнал мог быть не подключен
                pass

    # === НОВОЕ: Слот для обработки сигнала о предстоящем изменении ===
    @Slot(QModelIndex, object, object)
    def _on_cell_data_about_to_change(self, index: QModelIndex, old_value, new_value):
        """
        Слот, вызываемый, когда модель сигнализирует о предстоящем изменении данных.
        Здесь формируется EditAction и добавляется в стек Undo.
        """
        if not index.isValid():
            return
        row = index.row()
        col = index.column()
        logger.debug(f"SheetEditor: Получено уведомление об изменении [{row},{col}]: '{old_value}' -> '{new_value}'")
        action = EditAction(row=row, col=col, old_value=old_value, new_value=new_value)
        self._push_to_undo_stack(action)

    @Slot(QModelIndex, QModelIndex)
    def _on_model_data_changed(self, top_left: QModelIndex, bottom_right: QModelIndex):
        """
        Слот, вызываемый, когда модель сигнализирует об изменении данных.
        Здесь происходит вызов AppController для сохранения изменений.
        """
        logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed ВХОД")
        logger.debug(
            f"SheetEditor._on_model_data_changed вызван для диапазона: ({top_left.row()},{top_left.column()}) - ({bottom_right.row()},{bottom_right.column()})")
        if not self.app_controller or not self.sheet_name or not self._model:
            logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed ПРЕДУСЛОВИЯ НЕ ВЫПОЛНЕНЫ")
            logger.error("SheetEditor: Не хватает компонентов для обработки изменений.")
            return
        if not hasattr(self.app_controller, 'update_sheet_cell_in_project'):
            logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed НЕТ МЕТОДА")
            logger.error("SheetEditor: AppController не имеет метода 'update_sheet_cell_in_project'.")
            return
        logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed ПРЕДУСЛОВИЯ ПРОЙДЕНЫ")
        # Обрабатываем только одиночные ячейки
        if top_left == bottom_right and top_left.isValid():
            logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed ОБРАБОТКА ЯЧЕЙКИ")
            row = top_left.row()
            col = top_left.column()
            if not (0 <= row < self._model.rowCount() and 0 <= col < self._model.columnCount()):
                logger.warning(f"SheetEditor: Изменение за пределами модели. row={row}, col={col}")
                return
            column_name = self._model._original_column_names[
                col] if col < len(self._model._original_column_names) else f"Col_{col}"
            new_value = self._model._rows[row][
                col] if row < len(self._model._rows) and col < len(self._model._rows[row]) else None
            logger.debug(
                f"SheetEditor: Обнаружено изменение в ячейке [{row}, {column_name}]. Новое значение: '{new_value}'")
            try:
                logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed ВЫЗОВ КОНТРОЛЛЕРА")
                success = self.app_controller.update_sheet_cell_in_project(
                    sheet_name=self.sheet_name,
                    row_index=row,
                    column_name=column_name,
                    new_value=str(new_value) if new_value is not None else ""
                )
                logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed КОНТРОЛЛЕР ВЕРНУЛ: {success}")
                if success:
                    logger.info(
                        f"Изменение в ячейке [{self.sheet_name}][{row}, {column_name}] успешно сохранено в БД и истории.")
                else:
                    logger.error(
                        f"Не удалось сохранить изменение в ячейке [{self.sheet_name}][{row}, {column_name}] в БД.")
                    QMessageBox.warning(self, "Ошибка сохранения",
                                        f"Не удалось сохранить изменение в ячейке {column_name}{row + 1}.\n"
                                        "Изменение в интерфейсе будет сохранено до закрытия, но не записано в проект.")
            except Exception as e:
                logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed ИСКЛЮЧЕНИЕ: {e}")
                critical_error_msg = f"Исключение при сохранении изменения в ячейке [{self.sheet_name}][{row}, {column_name}]: {e}"
                logger.error(critical_error_msg, exc_info=True)
                QMessageBox.critical(self, "Критическая ошибка сохранения",
                                     f"Произошла ошибка при попытке сохранить изменение:\n{e}\n\n"
                                     "Изменение в интерфейсе будет сохранено до закрытия, но не записано в проект.")
        else:
            logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed НЕ ЯЧЕЙКА ИЛИ НЕ ВАЛИДНА")
            range_info = f"от ({top_left.row()},{top_left.column()}) до ({bottom_right.row()},{bottom_right.column()})"
            logger.debug(f"SheetEditor: Обнаружен диапазон изменений {range_info}. Обработка диапазонов не реализована.")
        logger.debug(f"DEBUG_SHEET_EDITOR: _on_model_data_changed ВЫХОД")

    # === Методы Undo/Redo ===
    def undo(self):
        """Отменяет последнее действие редактирования."""
        if not self._undo_stack or not self._model:
            logger.debug("SheetEditor: Нечего отменять или модель не доступна.")
            return
        action = self._undo_stack.pop()
        logger.debug(
            f"SheetEditor: Отмена действия: [{action.row}, {action.col}] '{action.new_value}' -> '{action.old_value}'")
        self._model.setDataInternal(action.row, action.col, action.old_value)
        self._redo_stack.append(action)
        if len(self._redo_stack) > self._max_undo_steps:
            self._redo_stack.pop(0)
        self._update_undo_redo_actions()
        logger.info(f"Отменено изменение в ячейке [{action.row}, {action.col}].")

    def redo(self):
        """Повторяет последнее отмененное действие."""
        if not self._redo_stack or not self._model:
            logger.debug("SheetEditor: Нечего повторять или модель не доступна.")
            return
        action = self._redo_stack.pop()
        logger.debug(
            f"SheetEditor: Повтор действия: [{action.row}, {action.col}] '{action.old_value}' -> '{action.new_value}'")
        self._model.setDataInternal(action.row, action.col, action.new_value)
        self._undo_stack.append(action)
        if len(self._undo_stack) > self._max_undo_steps:
            self._undo_stack.pop(0)
        self._update_undo_redo_actions()
        logger.info(f"Повторено изменение в ячейке [{action.row}, {action.col}].")

    def _push_to_undo_stack(self, action: EditAction):
        """Добавляет действие в стек отмены и очищает стек повторов."""
        self._undo_stack.append(action)
        if len(self._undo_stack) > self._max_undo_steps:
            self._undo_stack.pop(0)
        self._redo_stack.clear()
        self._update_undo_redo_actions()
        logger.debug(f"Действие добавлено в стек Undo: {action}")

    def _clear_undo_redo_stacks(self):
        """Очищает стеки Undo и Redo."""
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._update_undo_redo_actions()
        logger.debug("Стеки Undo/Redo очищены.")

    def _update_undo_redo_actions(self):
        """Обновляет состояние действий Undo/Redo."""
        self.action_undo.setEnabled(len(self._undo_stack) > 0)
        self.action_redo.setEnabled(len(self._redo_stack) > 0)

    def clear_sheet(self):
        logger.debug("Очистка редактора листа")
        self.project_db_path = None
        self.sheet_name = None
        self._clear_undo_redo_stacks()
        self.label_sheet_name.setText("Лист: <Не выбран>")
        self.table_view.setModel(None) # Это должно вызвать selectionModelChanged
        self._model = None
        # === НОВОЕ: Очистка строки редактирования ===
        self._clear_formula_bar() # Используем общий метод очистки
        # === НОВОЕ: Отключаем сигналы модели выделения при очистке ===
        self._disconnect_selection_model_signals()
        # =========================================================
        # =========================================
