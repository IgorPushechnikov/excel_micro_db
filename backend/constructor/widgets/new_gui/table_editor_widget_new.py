# backend/constructor/widgets/new_gui/table_editor_widget_new.py
"""
Виджет для редактирования листа Excel с помощью QTableView и строки формул.
Использует TableModel для отображения и изменения данных.
"""

import logging
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableView, QLineEdit, 
    QMessageBox, QHeaderView, QAbstractItemView
)
from PySide6.QtCore import Qt, QModelIndex, QItemSelection, QItemSelectionModel
from PySide6.QtGui import QKeySequence

# Импортируем модель
from .table_model import TableModel
# Импортируем AppController
from backend.core.app_controller import create_app_controller
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class TableEditorWidget(QWidget):
    """
    Виджет редактирования листа Excel.
    """

    def __init__(self, app_controller, parent=None):
        """
        Инициализирует виджет.

        Args:
            app_controller: Экземпляр AppController.
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self.app_controller = app_controller

        # --- Атрибуты виджета ---
        self.model: Optional[TableModel] = None
        self.table_view: Optional[QTableView] = None
        self.formula_line_edit: Optional[QLineEdit] = None
        
        # Текущий индекс для отслеживания изменений
        self._current_index: Optional[QModelIndex] = None
        # -----------------------

        self._setup_ui()
        self._setup_connections()

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # --- Создание строки формул ---
        formula_layout = QHBoxLayout()
        formula_layout.setContentsMargins(0, 0, 0, 0)
        self.formula_line_edit = QLineEdit(self)
        self.formula_line_edit.setPlaceholderText("Формула или значение...")
        formula_layout.addWidget(self.formula_line_edit)
        main_layout.addLayout(formula_layout)
        # -----------------------------

        # --- Создание таблицы ---
        self.table_view = QTableView(self)
        # Настройки таблицы
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.ContiguousSelection)
        self.table_view.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked |
            QAbstractItemView.EditTrigger.EditKeyPressed
        )
        
        # Настройка заголовков
        horizontal_header = self.table_view.horizontalHeader()
        horizontal_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        vertical_header = self.table_view.verticalHeader()
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        vertical_header.setDefaultSectionSize(20) # Высота строк по умолчанию
        
        main_layout.addWidget(self.table_view)
        # --------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        # --- Подключение сигналов таблицы ---
        if self.table_view:
            self.table_view.clicked.connect(self._on_cell_clicked)
            self.table_view.doubleClicked.connect(self._on_cell_double_clicked)
            selection_model = self.table_view.selectionModel()
            if selection_model:
                selection_model.selectionChanged.connect(self._on_selection_changed)
        # ----------------------------------

        # --- Подключение сигналов строки формул ---
        if self.formula_line_edit:
            self.formula_line_edit.returnPressed.connect(self._on_formula_return_pressed)
            self.formula_line_edit.editingFinished.connect(self._on_formula_editing_finished)
        # -----------------------------------------

    def load_sheet(self, sheet_name: str):
        """
        Загружает данные листа в таблицу.

        Args:
            sheet_name (str): Имя листа для загрузки.
        """
        logger.info(f"TableEditorWidget: Загрузка листа '{sheet_name}'...")
        try:
            if not self.app_controller or not self.app_controller.is_project_loaded:
                logger.error("Проект не загружен в AppController. Невозможно загрузить лист.")
                QMessageBox.critical(self, "Ошибка", "Проект не загружен. Невозможно загрузить лист.")
                return

            # Создаем или получаем модель
            if self.model is None and self.table_view:
                self.model = TableModel(self.app_controller, self)
                self.table_view.setModel(self.model)
                logger.debug("TableModel создана и установлена для QTableView.")
            
            # Загружаем данные в модель ТОЛЬКО ЕСЛИ МОДЕЛЬ СОЗДАНА
            if self.model: # <-- НОВАЯ ПРОВЕРКА
                self.model.load_sheet(sheet_name)
                logger.info(f"TableEditorWidget: Лист '{sheet_name}' успешно загружен в модель.")
            else:
                logger.error("TableModel не была создана или QTableView не инициализирован. Невозможно загрузить данные.")
                QMessageBox.critical(self, "Ошибка", "Внутренняя ошибка: модель таблицы не инициализирована.")
            
            # Сбрасываем текущий индекс
            self._current_index = None
            if self.formula_line_edit:
                assert self.formula_line_edit is not None  # <-- Удовлетворяет Pylance
                self.formula_line_edit.clear()
            
        except Exception as e:
            logger.error(f"Ошибка при загрузке листа '{sheet_name}' в TableEditorWidget: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке листа '{sheet_name}':\n{e}")

    # --- Обработчики событий таблицы ---
    def _on_cell_clicked(self, index: QModelIndex):
        """
        Обработчик клика по ячейке в QTableView.
        Обновляет строку формул значением из ячейки.
        """
        if index.isValid():
            self._current_index = index
            value = None
            if self.model:
                value = self.model.data(index, Qt.ItemDataRole.EditRole)
            if self.formula_line_edit:
                self.formula_line_edit.setText(str(value) if value is not None else "")
            logger.debug(f"Ячейка ({index.row()}, {index.column()}) выбрана. Значение: {value}")

    def _on_cell_double_clicked(self, index: QModelIndex):
        """
        Обработчик двойного клика по ячейке.
        Переводит ячейку в режим редактирования и фокусируется на строке формул.
        """
        if index.isValid():
            self._current_index = index
            # Фокус на строку формул
            if self.formula_line_edit:
                self.formula_line_edit.setFocus()
                # Выделяем текст для удобства
                self.formula_line_edit.selectAll()
            logger.debug(f"Двойной клик по ячейке ({index.row()}, {index.column()}). Фокус на формуле.")

    def _on_selection_changed(self, selected: QItemSelection, deselected: QItemSelection):
        """
        Обработчик изменения выделения в таблице.
        Если выделена одна ячейка, обновляет строку формул.
        """
        selected_indexes = selected.indexes()
        if len(selected_indexes) == 1:
            self._on_cell_clicked(selected_indexes[0])
        elif len(selected_indexes) > 1:
            # Если выдело несколько ячеек, можно очистить строку формул или показать что-то другое
            self._current_index = None
            if self.formula_line_edit: # <-- Добавляем assert и здесь тоже для единообразия, если Pylance ругается
                assert self.formula_line_edit is not None
                self.formula_line_edit.clear()
            logger.debug(f"Выдело {len(selected_indexes)} ячеек. Строка формул очищена.")
        # Если ничего не выделено, оставляем как есть или очищаем

    # --- Обработчики событий строки формул ---
    def _on_formula_return_pressed(self):
        """
        Обработчик нажатия Enter в строке формул.
        Сохраняет значение из строки формул в ячейку.
        """
        self._commit_formula()

    def _on_formula_editing_finished(self):
        """
        Обработчик завершения редактирования строки формул.
        (Например, потеря фокуса)
        """
        # В Qt editingFinished срабатывает при потере фокуса или нажатии Enter.
        # Так как Enter обрабатывается в _on_formula_return_pressed, здесь можно оставить пустым
        # или добавить логику, если нужно что-то делать при потере фокуса без Enter.
        # Например, можно спросить пользователя, хочет ли он сохранить изменения.
        # Пока оставим без действия, чтобы не было двойного сохранения.
        pass
        
    def _commit_formula(self):
        """
        Сохраняет значение из строки формул в текущую ячейку.
        """
        if self._current_index and self._current_index.isValid() and self.model and self.formula_line_edit:
            new_value = self.formula_line_edit.text()
            logger.debug(f"Сохранение значения '{new_value}' в ячейку ({self._current_index.row()}, {self._current_index.column()})")
            
            # setData модели обновит и данные, и вызовет dataChanged
            success = self.model.setData(self._current_index, new_value, Qt.ItemDataRole.EditRole)
            if success:
                logger.info(f"Значение ячейки ({self._current_index.row()}, {self._current_index.column()}) успешно обновлено через строку формул.")
                # Фокус возвращаем в таблицу
                if self.table_view:
                    self.table_view.setFocus()
            else:
                logger.error(f"Не удалось обновить значение ячейки ({self._current_index.row()}, {self._current_index.column()}) через строку формул.")
                # QMessageBox.critical(self, "Ошибка", f"Не удалось обновить значение ячейки ({self._current_index.row()+1}, {_index_to_column_name(self._current_index.column())}).")
        else:
            logger.warning("Нет активной ячейки для сохранения значения из строки формул.")
            
    # --- Вспомогательные функции ---
    def get_current_cell_address(self) -> Optional[str]:
        """
        Получает адрес текущей выделенной ячейки.
        
        Returns:
            Optional[str]: Адрес ячейки (например, "A1") или None.
        """
        if self._current_index and self._current_index.isValid() and self.model:
            row = self._current_index.row()
            col = self._current_index.column()
            return self.model.get_cell_address(row, col)
        return None
        
    def get_current_sheet_name(self) -> Optional[str]:
        """
        Получает имя текущего загруженного листа.
        
        Returns:
            Optional[str]: Имя листа или None.
        """
        if self.model:
            return self.model.get_sheet_name()
        return None
    # -----------------------------

# Вспомогательная функция для преобразования индекса в имя столбца (копия из table_model.py для локального использования)
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
