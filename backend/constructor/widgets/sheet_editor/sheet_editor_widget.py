# backend/constructor/widgets/sheet_editor/sheet_editor_widget.py
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
from backend.utils.logger import get_logger # <-- ИСПРАВЛЕНО: было from utils.logger

# --- ИМПОРТ МОДЕЛИ ДАННЫХ ИЗ ТОЙ ЖЕ ПОДПАПКИ ---
from .sheet_data_model import SheetDataModel
# -----------------------------------------------

logger = get_logger(__name__)


# === НОВОЕ: Структура для хранения информации об одном редактировании ===
class EditAction(NamedTuple):
    """Представляет одно действие редактирования для Undo/Redo."""
    row: int
    col: int
    old_value: Any
    new_value: Any


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
            # ИЗМЕНЕНО: Используем get_sheet_raw_data вместо get_sheet_editable_data
            # Это должно вернуть ВСЕ строки, включая строку заголовков.
            editable_data = self.app_controller.get_sheet_raw_data(sheet_name)
            if editable_data is not None and 'rows' in editable_data: # Проверяем наличие данных
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
                            f"Столбцов: {len(editable_data.get('rows', [])[0]) if editable_data.get('rows') else 0}")

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
                        # --- ДОБАВЛЕНО ЛОГИРОВАНИЕ ДЛЯ ОТЛАДКИ ---
                        if styles_data:
                            logger.debug(f"SheetEditor.load_sheet: Пример стиля: {styles_data[0]}")
                        # ======================================
                        # Передаем стили в модель
                        self._model.set_cell_styles(styles_data)
                        # --- ДОБАВЛЕНО ЛОГИРОВАНИЕ ДЛЯ ОТЛАДКИ ---
                        logger.debug(f"SheetEditor.load_sheet: Стили переданы в модель для листа '{sheet_name}'.")
                        # ======================================

                        # === НОВОЕ: Явный вызов dataChanged для всей модели ===
                        # Это гарантирует, что представление обновится, даже если сигналы не сработали должным образом.
                        if self._model.rowCount() > 0 and self._model.columnCount() > 0:
                            top_left = self._model.index(0, 0)
                            bottom_right = self._model.index(self._model.rowCount() - 1, self._model.columnCount() - 1)
                            if top_left.isValid() and bottom_right.isValid():
                                self._model.dataChanged.emit(top_left, bottom_right)
                                logger.debug("SheetEditor.load_sheet: Явный сигнал dataChanged отправлен для всей модели.")
                        # ===================================================

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
            # ИСПРАВЛЕНО: Используем индекс столбца как имя, так как _original_column_names больше нет
            column_name = str(col) # Используем индекс столбца как имя
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
