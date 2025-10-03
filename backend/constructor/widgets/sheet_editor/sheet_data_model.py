# backend/constructor/widgets/sheet_editor/sheet_data_model.py
"""
Модель данных для отображения и редактирования содержимого листа Excel в QTableView.
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

# --- ИМПОРТ МОДУЛЯ ФОРМАТИРОВАНИЯ ИЗ ТОЙ ЖЕ ПОДПАПКИ ---
from .cell_formatting import format_cell_value
# ------------------------------------------------------

logger = get_logger(__name__)


# =====================================================================
# === SheetDataModel с исправленными сигнатурами и логикой ===
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
        # Логирование инициализации модели
        logger.debug(f"SheetDataModel.__init__: _rows count = {len(self._rows)}, column count = {len(self._generated_column_headers)}")

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
                col_name_for_tooltip = self._generated_column_headers[col] if col < len(self._generated_column_headers) else f"Col_{col}"
                return f"Столбец: {col_name_for_tooltip}\nЗначение: {repr(value)}"
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
        logger.debug(f"SheetDataModel.headerData: section={section}, orientation={orientation}, role={role}")
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                # Возвращаем стандартные имена Excel (A, B, C...), как и было изначально
                if 0 <= section < len(self._generated_column_headers):
                    header_val = self._generated_column_headers[section]
                    logger.debug(f"SheetDataModel.headerData: Returning Excel-style column header '{header_val}' for section {section}")
                    return header_val
                else:
                    fallback_header = f"Col_{section}"
                    logger.debug(f"SheetDataModel.headerData: Returning fallback column header '{fallback_header}' for section {section}")
                    return fallback_header
            elif orientation == Qt.Orientation.Vertical:
                # Номера строк (1-based), как в Excel
                row_header = str(section + 1)
                logger.debug(f"SheetDataModel.headerData: Returning row header '{row_header}' for section {section}")
                return row_header
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