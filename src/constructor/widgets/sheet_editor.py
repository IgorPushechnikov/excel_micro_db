# src/constructor/widgets/sheet_editor.py
"""
Виджет-редактор для отображения и редактирования содержимого листа Excel.
"""

import sys
import string  # Для генерации имен столбцов Excel

from typing import Optional, Dict, Any, List, NamedTuple, Union
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTableView, QLabel, QMessageBox,
    QAbstractItemView, QHeaderView, QApplication, QMenu, QInputDialog, QHBoxLayout, QLineEdit
)
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, Slot, Signal # Импортируем QModelIndex
from PySide6.QtGui import QBrush, QColor, QAction # Импортируем QAction

import sqlite3
import logging

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from src.core.app_controller import AppController

from src.utils.logger import get_logger

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
        for style_info in styles_data:
            range_addr = style_info.get("range_address", "")
            style_attrs = style_info.get("style_attributes", {})
            if range_addr and style_attrs:
                # Очень упрощенный парсер адреса ячейки
                try:
                    col_letter = ''.join(filter(str.isalpha, range_addr)).upper()
                    row_number_str = ''.join(filter(str.isdigit, range_addr))
                    if col_letter and row_number_str:
                        col_index = self._column_letter_to_index(col_letter)
                        row_index = int(row_number_str) - 1  # Excel 1-based -> Python 0-based
                        if col_index >= 0 and row_index >= 0:
                            self._cell_styles[(row_index, col_index)] = style_attrs
                except Exception as e:
                    logger.warning(f"Не удалось распарсить адрес ячейки '{range_addr}' для стиля: {e}")

        # Сообщаем представлению, что данные могли измениться (для перерисовки стилей)
        top_left = self.index(0, 0)
        bottom_right = self.index(self.rowCount() - 1, self.columnCount() - 1)
        if top_left.isValid() and bottom_right.isValid():
            self.dataChanged.emit(top_left, bottom_right)

    def _column_letter_to_index(self, letter: str) -> int:
        """Преобразует букву столбца Excel (например, 'A', 'Z', 'AA') в 0-базовый индекс."""
        result = 0
        for char in letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1  # 0-based index

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

    def rowCount(self, parent=QModelIndex()) -> int: # Убрана аннотация Union
        _ = parent # Чтобы избежать предупреждения о неиспользованном параметре
        return len(self._rows)

    def columnCount(self, parent=QModelIndex()) -> int: # Убрана аннотация Union
        _ = parent # Чтобы избежать предупреждения о неиспользованном параметре
        return len(self._rows[0]) if self._rows else 0

    # ИСПРАВЛЕНО: Сигнатура метода data соответствует базовому классу
    def data(self, index, role: int = Qt.ItemDataRole.DisplayRole): # role: int
        # Проверка валидности индекса
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if role == Qt.ItemDataRole.DisplayRole:
            if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
                value = self._rows[row][col]
                return str(value) if value is not None else ""
        elif role == Qt.ItemDataRole.ToolTipRole:
            if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
                value = self._rows[row][col]
                orig_name = self._original_column_names[col] if col < len(self._original_column_names) else f"Col_{col}"
                return f"Столбец: {orig_name}\nЗначение: {repr(value)}"
        # === НОВОЕ: Обработка ролей для стилей ===
        elif role == Qt.ItemDataRole.BackgroundRole:
            style = self._cell_styles.get((row, col))
            if style:
                fill_fg_color = style.get("fill_fg_color")
                if fill_fg_color and isinstance(fill_fg_color, str) and fill_fg_color.startswith('#'):
                    try:
                        color = QColor(fill_fg_color)
                        if color.isValid():
                            return QBrush(color)
                    except Exception as e:
                        logger.warning(f"Ошибка преобразования цвета фона '{fill_fg_color}': {e}")
        elif role == Qt.ItemDataRole.ForegroundRole:
            style = self._cell_styles.get((row, col))
            if style:
                font_color = style.get("font_color")
                if font_color and isinstance(font_color, str) and font_color.startswith('#'):
                    try:
                        color = QColor(font_color)
                        if color.isValid():
                            return QBrush(color)
                    except Exception as e:
                        logger.warning(f"Ошибка преобразования цвета шрифта '{font_color}': {e}")
        # ======================================

        return None

    # ИСПРАВЛЕНО: Сигнатура метода headerData соответствует базовому классу
    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole): # role: int
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                if 0 <= section < len(self._generated_column_headers):
                    return self._generated_column_headers[section]
                else:
                    return f"Col_{section}"  # fallback
            elif orientation == Qt.Orientation.Vertical:
                return str(section + 1) # Номера строк (1-based)
        return None

    # ИСПРАВЛЕНО: Сигнатура метода flags соответствует базовому классу
    def flags(self, index) -> Qt.ItemFlag: # Убрана аннотация Union
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled

    # ИСПРАВЛЕНО: Сигнатура метода setData соответствует базовому классу
    def setData(self, index, value, role: int = Qt.ItemDataRole.EditRole): # role: int
        """
        Устанавливает данные в модель. Вызывается, когда пользователь редактирует ячейку.
        Испускает cellDataAboutToChange до изменения и dataChanged после.
        """
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
                self.dataChanged.emit(index, index, [role]) # role передается как список
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


# === SheetEditor с исправлениями для Pylance и логики ===
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

        # === НОВОЕ: Строка редактирования ===
        self._formula_bar: Optional[QLineEdit] = None
        self._current_editing_index: Optional[QModelIndex] = None
        # ===================================

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.label_sheet_name = QLabel("Лист: <Не выбран>")
        self.label_sheet_name.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(self.label_sheet_name)

        # === НОВОЕ: Добавление строки редактирования ===
        formula_layout = QHBoxLayout()
        formula_label = QLabel("Формула/Значение:")
        self._formula_bar = QLineEdit()
        # Проверка на None не обязательна здесь, так как мы только что создали объект
        if self._formula_bar:
             self._formula_bar.setPlaceholderText("Выберите ячейку или введите значение/формулу")
             self._formula_bar.returnPressed.connect(self._on_formula_bar_return_pressed)
        formula_layout.addWidget(formula_label)
        formula_layout.addWidget(self._formula_bar)
        layout.addLayout(formula_layout)
        # =============================================

        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.ContiguousSelection)

        # === НОВОЕ: Подключение сигнала выбора ячейки ===
        self.table_view.clicked.connect(self._on_cell_clicked)
        # =============================================

        # Настройка контекстного меню
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self._on_context_menu)

        horizontal_header = self.table_view.horizontalHeader()
        horizontal_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        vertical_header = self.table_view.verticalHeader()
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        vertical_header.setDefaultSectionSize(20)

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
        if not index.isValid() or not self._model:
            self._current_editing_index = None
            # Проверяем, что _formula_bar существует
            if self._formula_bar:
                self._formula_bar.setText("")
                self._formula_bar.setPlaceholderText("Выберите ячейку или введите значение/формулу")
            return

        self._current_editing_index = index
        # Получаем значение ячейки из модели для отображения в строке редактирования
        display_value = self._model.data(index, Qt.ItemDataRole.DisplayRole)
        # Проверяем, что _formula_bar существует
        if self._formula_bar:
            # Убеждаемся, что display_value - строка
            self._formula_bar.setText(display_value if isinstance(display_value, str) else str(display_value) if display_value is not None else "")

    @Slot()
    def _on_formula_bar_return_pressed(self):
        """Обработчик нажатия Enter в строке редактирования."""
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
        self._current_editing_index = None
        if self._formula_bar:
            self._formula_bar.setText("")
            self._formula_bar.setPlaceholderText("Выберите ячейку или введите значение/формулу")
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
                self.table_view.setModel(self._model)
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
                        from src.storage.styles import load_sheet_styles
                        styles_data = load_sheet_styles(conn, sheet_id)
                        logger.debug(f"Загружено {len(styles_data)} стилей для листа '{sheet_name}'.")
                        # Передаем стили в модель
                        self._model.set_cell_styles(styles_data)
                    else:
                        logger.warning(f"Лист '{sheet_name}' не найден в БД при попытке загрузить стили.")
                    conn.close()
                except Exception as e:
                    logger.error(f"Ошибка при загрузке стилей для листа '{sheet_name}': {e}")
                # =================================================

            else:
                self.table_view.setModel(None)
                self._model = None
                logger.warning(f"Редактируемые данные для листа '{sheet_name}' не найдены или пусты.")
        except Exception as e:
            logger.error(f"Ошибка при загрузке листа '{sheet_name}': {e}", exc_info=True)
            QMessageBox.critical(
                self,
                "Ошибка загрузки",
                f"Не удалось загрузить содержимое листа '{sheet_name}':\n{e}"
            )
            self.table_view.setModel(None)
            self._model = None

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
        self.table_view.setModel(None)
        self._model = None
        # === НОВОЕ: Очистка строки редактирования ===
        self._current_editing_index = None
        if self._formula_bar:
             self._formula_bar.setText("")
             self._formula_bar.setPlaceholderText("Выберите ячейку или введите значение/формулу")
        # =========================================
