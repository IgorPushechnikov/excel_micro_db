# backend/constructor/widgets/new_gui/table_editor_widget_new.py
"""
Виджет для редактирования листа Excel с помощью QTableView и строки формул.
Использует TableModel для отображения и изменения данных.
"""

import logging
import re
import csv # <-- Добавлен импорт csv
import io # <-- Добавлен импорт io
from enum import Enum
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableView, QLineEdit, 
    QMessageBox, QHeaderView, QAbstractItemView, QToolBar, QPushButton,
    QInputDialog, QStatusBar, QApplication, QMenu # <-- Добавлен QMenu
)

from PySide6.QtCore import Qt, QModelIndex, QItemSelection, QItemSelectionModel
from PySide6.QtGui import QKeySequence, QCursor, QGuiApplication # <-- Добавлен QGuiApplication

# Импортируем модель
from .table_model import TableModel
# Импортируем AppController
from backend.core.app_controller import create_app_controller
from backend.utils.logger import get_logger
# Импортируем калькулятор формул
from backend.core.formula_calculators import apply_age_formula_to_column
# --- ИМПОРТ НОВОЙ УТИЛИТЫ ---
from backend.utils.ui_loader import load_context_menu_from_yaml, UICommandHandler # <-- Добавлен импорт
# --- КОНЕЦ ИМПОРТА ---

logger = get_logger(__name__)

# --- НОВОЕ: Перечисление для состояния выбора формулы возраста ---
class AgeFormulaSelectionStep(Enum):
    WAITING_START_DATE = 0
    WAITING_END_DATE = 1
    WAITING_RESULT_COLUMN = 2
    NOT_ACTIVE = -1
# --- КОНЕЦ НОВОГО ---

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
        # --- НОВОЕ: Атрибуты для тулбара ---
        self.toolbar: Optional[QToolBar] = None
        self.apply_age_formula_button: Optional[QPushButton] = None
        # --- КОНЕЦ НОВОГО ---
        
        # --- НОВОЕ: Атрибуты для режима выбора формулы возраста ---
        self._waiting_for_age_formula_input = False
        self._age_formula_step = AgeFormulaSelectionStep.NOT_ACTIVE
        self._age_formula_start_addr: Optional[str] = None
        self._age_formula_end_addr: Optional[str] = None
        self._age_formula_result_col: Optional[str] = None
        # --- КОНЕЦ НОВОГО ---

        # --- НОВОЕ: Атрибут command_handler ---
        self.command_handler: Optional[UICommandHandler] = None
        # --- КОНЕЦ НОВОГО ---
        
        # Текущий индекс для отслеживания изменений
        self._current_index: Optional[QModelIndex] = None
        # -----------------------

        self._setup_ui()
        self._setup_connections()
        # --- ИНИЦИАЛИЗАЦИЯ command_handler ---
        self.command_handler = UICommandHandler(self)
        self.command_handler.register_handler("_on_copy_triggered", self._on_copy_triggered)
        self.command_handler.register_handler("_on_paste_triggered", self._on_paste_triggered)
        # --- КОНЕЦ ИНИЦИАЛИЗАЦИИ ---

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # --- НОВОЕ: Создание тулбара ---
        self.toolbar = QToolBar(self)
        self.toolbar.setMovable(False) # Закрепляем тулбар
        # Создаём кнопку
        self.apply_age_formula_button = QPushButton("Применить формулу возраста", self)
        self.toolbar.addWidget(self.apply_age_formula_button)
        main_layout.addWidget(self.toolbar)
        # --- КОНЕЦ НОВОГО ---

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
        # --- НОВОЕ: Подключение сигнала кнопки ---
        if self.apply_age_formula_button:
            self.apply_age_formula_button.clicked.connect(self._on_apply_age_formula_clicked)
        # --- КОНЕЦ НОВОГО ---

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

    # --- НОВОЕ: Обработчик нажатия кнопки ---
    def _on_apply_age_formula_clicked(self):
        """
        Обработчик нажатия кнопки "Применить формулу возраста".
        Переходит в режим выбора ячеек.
        """
        if not self.app_controller or not self.app_controller.is_project_loaded:
            QMessageBox.critical(self, "Ошибка", "Проект не загружен. Невозможно применить формулу.")
            return

        sheet_name = self.get_current_sheet_name()
        if not sheet_name:
            QMessageBox.critical(self, "Ошибка", "Нет загруженного листа. Невозможно применить формулу.")
            return

        # Начинаем режим выбора
        self._start_age_formula_selection()

    def _start_age_formula_selection(self):
        """Инициирует режим выбора ячеек для формулы возраста."""
        self._waiting_for_age_formula_input = True
        self._age_formula_step = AgeFormulaSelectionStep.WAITING_START_DATE
        self._age_formula_start_addr = None
        self._age_formula_end_addr = None
        self._age_formula_result_col = None
        
        # Изменяем курсор и подсказку (если есть статусная строка, можно туда)
        assert self.table_view is not None
        self.table_view.setCursor(QCursor(Qt.CursorShape.CrossCursor))
        assert self.formula_line_edit is not None
        self.formula_line_edit.setPlaceholderText("Выберите ячейку с начальной датой...")
        logger.info("Режим выбора формулы возраста активирован. Ожидание выбора начальной даты.")

    def _on_cell_clicked_for_age_formula(self, index: QModelIndex):
        """
        Обработка клика по ячейке в режиме выбора формулы возраста.
        """
        if not index.isValid() or not self.model:
            return

        cell_addr = self.model.get_cell_address(index.row(), index.column())
        if not cell_addr:
            return

        if self._age_formula_step == AgeFormulaSelectionStep.WAITING_START_DATE:
            self._age_formula_start_addr = cell_addr
            # Визуально выделяем ячейку (можно сбросить предыдущие выделения)
            if self.table_view and self.table_view.selectionModel():
                selection_model = self.table_view.selectionModel()
                selection_model.clearSelection()
                selection_model.select(
                    QItemSelection(index, index),
                    QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows
                )
            assert self.formula_line_edit is not None
            self.formula_line_edit.setPlaceholderText(f"Нач. дата: {cell_addr}. Выберите ячейку с конечной датой...")
            self._age_formula_step = AgeFormulaSelectionStep.WAITING_END_DATE
            logger.debug(f"Выбрана ячейка начальной даты: {cell_addr}")

        elif self._age_formula_step == AgeFormulaSelectionStep.WAITING_END_DATE:
            self._age_formula_end_addr = cell_addr
            # Визуально выделяем ячейку
            if self.table_view and self.table_view.selectionModel():
                selection_model = self.table_view.selectionModel()
                # Не очищаем, чтобы видеть обе ячейки
                selection_model.select(
                    QItemSelection(index, index),
                    QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows
                )
            # Извлекаем букву столбца из адреса
            result_col = "".join(filter(str.isalpha, cell_addr))
            assert self.formula_line_edit is not None
            self.formula_line_edit.setPlaceholderText(f"Кон. дата: {cell_addr}. Выберите столбец результата (например, кликните на заголовок или ячейку в столбце {result_col})...")
            self._age_formula_step = AgeFormulaSelectionStep.WAITING_RESULT_COLUMN
            logger.debug(f"Выбрана ячейка конечной даты: {cell_addr}")

        elif self._age_formula_step == AgeFormulaSelectionStep.WAITING_RESULT_COLUMN:
            # Извлекаем букву столбца из адреса
            result_col = "".join(filter(str.isalpha, cell_addr)).upper()
            self._age_formula_result_col = result_col
            # Визуально выделяем ячейку (опционально, можно не выделять столбец, а только сообщить)
            # selection_model = self.table_view.selectionModel()
            # selection_model.select(QItemSelection(index, index), QItemSelectionModel.Select)
            assert self.formula_line_edit is not None
            self.formula_line_edit.setPlaceholderText(f"Столбец результата: {result_col}. Применяем формулу...")
            logger.debug(f"Выбран результирующий столбец: {result_col}")
            # Завершаем режим и вызываем калькулятор
            self._finish_age_formula_selection()

    def _finish_age_formula_selection(self):
        """
        Завершает режим выбора и вызывает калькулятор.
        """
        # Проверяем, все ли адреса собраны
        if self._age_formula_start_addr and self._age_formula_end_addr and self._age_formula_result_col:
            sheet_name = self.get_current_sheet_name()
            if sheet_name:
                # Вызываем функцию из formula_calculators
                success = apply_age_formula_to_column(
                    self.app_controller.data_manager,
                    sheet_name,
                    self._age_formula_start_addr,
                    self._age_formula_end_addr,
                    self._age_formula_result_col
                )
                if success:
                    QMessageBox.information(self, "Успех", f"Формула возраста успешно применена к столбцу {self._age_formula_result_col}.")
                    # Перезагружаем данные, чтобы отобразить изменения
                    self.load_sheet(sheet_name)
                else:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось применить формулу возраста к столбцу {self._age_formula_result_col}.")
            else:
                QMessageBox.critical(self, "Ошибка", "Не удалось определить имя текущего листа.")
        else:
            logger.warning("Не все адреса были выбраны для формулы возраста.")
            QMessageBox.warning(self, "Предупреждение", "Не все адреса были выбраны. Операция отменена.")
        # Сбрасываем режим *после* проверки и выполнения
        self._cancel_age_formula_selection()

    def _cancel_age_formula_selection(self):
        """
        Сбрасывает режим выбора формулы возраста.
        """
        self._waiting_for_age_formula_input = False
        self._age_formula_step = AgeFormulaSelectionStep.NOT_ACTIVE
        self._age_formula_start_addr = None
        self._age_formula_end_addr = None
        self._age_formula_result_col = None
        
        assert self.table_view is not None
        self.table_view.setCursor(QCursor(Qt.CursorShape.ArrowCursor))
        assert self.formula_line_edit is not None
        self.formula_line_edit.setPlaceholderText("Формула или значение...")
        # Очищаем выделение, если нужно
        if self.table_view and self.table_view.selectionModel():
            self.table_view.selectionModel().clearSelection()
        logger.info("Режим выбора формулы возраста отменён.")
    # --- КОНЕЦ НОВОГО ---

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
        # --- НОВОЕ: Проверка режима выбора формулы возраста ---
        if self._waiting_for_age_formula_input:
            self._on_cell_clicked_for_age_formula(index)
            return # Не выполняем стандартную логику
        # --- КОНЕЦ НОВОГО ---
        
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
        # --- НОВОЕ: Игнорировать двойной клик в режиме выбора ---
        if self._waiting_for_age_formula_input:
            return
        # --- КОНЕЦ НОВОГО ---
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
        # --- НОВОЕ: Игнорировать изменение выделения в режиме выбора ---
        if self._waiting_for_age_formula_input:
            return
        # --- КОНЕЦ НОВОГО ---
        
        selected_indexes = selected.indexes()
        if len(selected_indexes) == 1:
            self._on_cell_clicked(selected_indexes[0])
        elif len(selected_indexes) > 1:
            # Если выдело нескольких ячеек, можно очистить строку формул или показать что-то другое
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
            
    # --- НОВОЕ: Обработка нажатия клавиш ---
    def keyPressEvent(self, event):
        """
        Обработка нажатия клавиш.
        Используется для отмены режима выбора по клавише Esc.
        """
        # Если активен режим выбора и нажата клавиша Esc (и строка формул не в фокусе, чтобы не мешать её очистке)
        if self._waiting_for_age_formula_input and event.key() == Qt.Key.Key_Escape and self.focusWidget() != self.formula_line_edit:
            self._cancel_age_formula_selection()
            event.accept() # Помечаем событие как обработанное
            return
        # Вызываем стандартную обработку
        super().keyPressEvent(event)
    # --- КОНЕЦ НОВОГО ---

    # --- НОВОЕ: Контекстное меню через YAML ---
    def contextMenuEvent(self, event):
        """
        Обработчик события контекстного меню.
        Загружает меню из YAML и показывает его.
        """
        # Путь к YAML-файлу (относительно корня проекта или можно передать как аргумент)
        yaml_path = "config/context_menus/table_editor.yaml" # Или использовать Path и app_controller.project_root

        # Загружаем меню
        menu = load_context_menu_from_yaml(yaml_path, self, self.command_handler)

        if menu:
            # Показываем меню в позиции курсора
            menu.exec_(event.globalPos())
        else:
            # Если меню не загрузилось, можно показать стандартное или просто игнорировать
            logger.warning("Контекстное меню не было загружено из YAML.")
            # super().contextMenuEvent(event) # Показать стандартное (если нужно)

    # --- КОНЕЦ КОНТЕКСТНОГО МЕНЮ ---

    # --- НОВОЕ: Метод копирования ---
    def _on_copy_triggered(self):
        """
        Обработчик команды 'Копировать'.
        Копирует выделенные ячейки (или все загруженные, если нет выделения) в буфер обмена как TSV.
        """
        logger.info("Команда 'Копировать' активирована.")
        try:
            # 1. Определить диапазон для копирования
            selected_indexes = self.table_view.selectionModel().selectedIndexes()
            if not selected_indexes:
                 # Если нет выделения, копируем все загруженные данные из модели
                 # Это зависит от того, как организована модель. Метод get_current_sheet_name() есть.
                 # Но как получить *все* данные из модели TableModel?
                 # TableModel хранит self._display_data. Но TableModel не предоставляет прямого метода
                 # для получения всех данных в "сыром" виде. Нужно адаптировать.
                 # Пока используем выделение. Если выделения нет, копируем пустую строку или выводим сообщение.
                 logger.info("Нет выделения. Копирование отменено.")
                 # Можно показать QMessageBox
                 # QMessageBox.information(self, "Копирование", "Нет выделения для копирования.")
                 return

            # 2. Найти границы выделения
            rows = [idx.row() for idx in selected_indexes]
            cols = [idx.column() for idx in selected_indexes]
            min_row, max_row = min(rows), max(rows)
            min_col, max_col = min(cols), max(cols)

            # 3. Получить значения из модели для этого диапазона
            tsv_data = []
            for r in range(min_row, max_row + 1):
                row_data = []
                for c in range(min_col, max_col + 1):
                    index = self.model.index(r, c)
                    value = self.model.data(index, Qt.ItemDataRole.DisplayRole) # Используем DisplayRole для копирования
                    # Обработка None или других типов для TSV
                    row_data.append(str(value) if value is not None else "")
                tsv_data.append(row_data)

            # 4. Преобразовать в TSV строку
            # Используем StringIO для создания строки
            output = io.StringIO()
            writer = csv.writer(output, delimiter='\t', lineterminator='\n') # TSV
            writer.writerows(tsv_data)
            tsv_string = output.getvalue()
            output.close() # Закрываем StringIO

            # 5. Поместить в буфер обмена
            clipboard = QGuiApplication.clipboard()
            clipboard.setText(tsv_string)

            logger.info(f"Скопировано {len(tsv_data)} строк в буфер обмена.")

        except Exception as e:
            logger.error(f"Ошибка при копировании в буфер обмена: {e}", exc_info=True)
            # Показать сообщение пользователю
            QMessageBox.critical(self, "Ошибка копирования", f"Произошла ошибка при копировании: {e}")

    # --- КОНЕЦ МЕТОДА КОПИРОВАТЬ ---

    # --- НОВОЕ: Метод вставки ---
    def _on_paste_triggered(self):
        """
        Обработчик команды 'Вставить'.
        Вставляет данные из буфера обмена (TSV) в БД, начиная с активной ячейки.
        """
        logger.info("Команда 'Вставить' активирована.")
        try:
            # 1. Получить строку из буфера обмена
            clipboard = QGuiApplication.clipboard()
            clipboard_text = clipboard.text()

            if not clipboard_text.strip(): # Проверяем на пустую строку или только пробельные символы
                logger.info("Буфер обмена пуст.")
                # Показать сообщение пользователю
                QMessageBox.information(self, "Вставка", "Буфер обмена пуст.")
                return

            # 2. Определить ячейку назначения (активная или A1)
            current_index = self.table_view.currentIndex()
            if current_index.isValid():
                start_row = current_index.row()
                start_col = current_index.column()
                logger.debug(f"Вставка начнётся с ячейки ({start_row}, {start_col}) (0-based).")
            else:
                start_row, start_col = 0, 0 # Если нет активной ячейки, начинаем с A1 (0, 0)
                logger.debug(f"Нет активной ячейки. Вставка начнётся с ячейки (0, 0).")

            # 3. Разобрать TSV строку в матрицу
            # Используем StringIO для чтения строки как файла
            input_stream = io.StringIO(clipboard_text)
            reader = csv.reader(input_stream, delimiter='\t') # TSV
            tsv_matrix = list(reader) # Получаем список списков
            input_stream.close() # Закрываем StringIO

            if not tsv_matrix or not any(row for row in tsv_matrix): # Проверяем, есть ли хоть одна непустая строка
                 logger.info("Буфер обмена содержит пустой TSV.")
                 QMessageBox.information(self, "Вставка", "Буфер обмена содержит пустой TSV.")
                 return

            logger.debug(f"Разобрано {len(tsv_matrix)} строк и {max(len(row) for row in tsv_matrix) if tsv_matrix else 0} столбцов из TSV.")

            # 4. Преобразовать матрицу в raw_data_list
            raw_data_list = []
            for r_idx, row in enumerate(tsv_matrix):
                for c_idx, value in enumerate(row):
                    target_row = start_row + r_idx
                    target_col = start_col + c_idx
                    cell_address = self.model.get_cell_address(target_row, target_col) # Используем существующий метод из модели
                    if cell_address: # Убедимся, что адрес сформирован корректно
                        raw_data_list.append({
                            "cell_address": cell_address, # e.g., 'A1'
                            "value": value # Значение из буфера
                        })

            if not raw_data_list:
                 logger.warning("Не удалось сформировать raw_data_list из TSV.")
                 QMessageBox.warning(self, "Вставка", "Не удалось обработать данные из буфера обмена.")
                 return

            # 5. Получить имя текущего листа
            sheet_name = self.get_current_sheet_name()
            if not sheet_name:
                logger.error("Не удалось определить имя текущего листа для вставки.")
                QMessageBox.critical(self, "Ошибка вставки", "Не удалось определить активный лист.")
                return

            # 6. Убедиться, что лист существует в БД (create или get sheet_id)
            # AppController.project_db_path должен быть доступен через self.app_controller
            # storage = self.app_controller.storage # <- Проверить, что storage инициализирован
            # storage.save_sheet(project_id=1, sheet_name=sheet_name) # <- Это обновит или создаст
            # Однако, AppController не предоставляет напрямую вызов save_sheet без sheet_id.
            # Но save_sheet_raw_data вызывает save_sheet внутри себя, если лист не существует.
            # Проверим, что AppController.storage подключен.
            if not self.app_controller.storage:
                 logger.error("AppController.storage не инициализирован.")
                 QMessageBox.critical(self, "Ошибка вставки", "Соединение с БД проекта не установлено.")
                 return

            # 7. Сохранить raw_data_list в БД
            success = self.app_controller.storage.save_sheet_raw_data(sheet_name, raw_data_list)
            if not success:
                 logger.error(f"Не удалось сохранить данные в БД для листа '{sheet_name}'.")
                 QMessageBox.critical(self, "Ошибка вставки", f"Не удалось сохранить данные в БД.")
                 return

            logger.info(f"Успешно вставлено {len(raw_data_list)} ячеек в лист '{sheet_name}' через БД.")

            # 8. Обновить GUI (перезагрузить данные модели)
            self.load_sheet(sheet_name) # <- Это обновит TableModel и QTableView

        except csv.Error as ce:
            logger.error(f"Ошибка разбора TSV из буфера обмена: {ce}", exc_info=True)
            QMessageBox.critical(self, "Ошибка вставки", f"Ошибка формата данных (TSV): {ce}")
        except Exception as e:
            logger.error(f"Ошибка при вставке из буфера обмена: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка вставки", f"Произошла ошибка при вставке: {e}")

    # --- КОНЕЦ МЕТОДА ВСТАВИТЬ ---

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
