# backend/constructor/widgets/new_gui/sheet_editor_widget.py
"""
Виджет для редактирования листа Excel с помощью QTableView и строки формул.
Использует DBTableModel для отображения и изменения данных.
"""

import logging
from PySide6.QtWidgets import QWidget, QVBoxLayout, QTableView, QLineEdit, QHBoxLayout, QPushButton
from PySide6.QtCore import Qt, QModelIndex
from PySide6.QtGui import QKeySequence, QShortcut, QClipboard, QApplication

from .qt_model_adapter import DBTableModel
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class SheetEditorWidget(QWidget):
    """
    Виджет редактирования листа Excel.
    """

    def __init__(self, app_controller, sheet_name: str, parent=None):
        """
        Инициализирует виджет.

        Args:
            app_controller: Экземпляр AppController.
            sheet_name (str): Имя листа для редактирования.
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self.app_controller = app_controller
        self.sheet_name = sheet_name

        self.model = DBTableModel(self.app_controller, self.sheet_name, self)
        self.table_view = QTableView(self)
        self.table_view.setModel(self.model)

        # Настройка строки формул
        self.formula_line_edit = QLineEdit(self)
        self.formula_line_edit.setPlaceholderText("Формула или значение...")

        # Кнопки
        self.commit_button = QPushButton("Ввод", self)
        self.revert_button = QPushButton("Отмена", self)

        # Макет
        main_layout = QVBoxLayout(self)

        # Макет для строки формул и кнопок
        formula_layout = QHBoxLayout()
        formula_layout.addWidget(self.formula_line_edit)
        formula_layout.addWidget(self.commit_button)
        formula_layout.addWidget(self.revert_button)

        main_layout.addLayout(formula_layout)
        main_layout.addWidget(self.table_view)

        # Подключение сигналов
        self.table_view.clicked.connect(self._on_cell_clicked)
        self.formula_line_edit.returnPressed.connect(self._on_commit_pressed)
        self.commit_button.clicked.connect(self._on_commit_pressed)
        self.revert_button.clicked.connect(self._on_revert_pressed)

        # Горячая клавиша F2 для редактирования
        self._f2_shortcut = QShortcut(QKeySequence(Qt.Key.Key_F2), self.table_view)
        self._f2_shortcut.activated.connect(self._on_f2_pressed)

        # Текущий индекс для отслеживания изменений
        self._current_index = QModelIndex()

    def keyPressEvent(self, event):
        """
        Обрабатывает события нажатия клавиш.
        Реализует вставку из буфера обмена по Ctrl+V.
        """
        # Проверяем, нажат ли Ctrl+V
        if event.key() == Qt.Key.Key_V and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self._paste_from_clipboard()
            event.accept() # Сообщаем, что событие обработано
            return
        # Вызываем базовую реализацию для остальных клавиш
        super().keyPressEvent(event)

    def _paste_from_clipboard(self):
        """
        Извлекает данные из буфера обмена и вставляет их в модель,
        начиная с текущего индекса.
        """
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text()
        if not clipboard_text:
            logger.debug("Буфер обмена пуст.")
            return

        logger.debug(f"Получены данные из буфера обмена (первые 100 символов): {clipboard_text[:100]}...")

        # Разбор данных из буфера (обычно табуляция и новая строка)
        rows = clipboard_text.strip().split('\n')
        if not rows:
            logger.debug("Нет строк для вставки из буфера.")
            return

        # Удаляем пустую строку, если она была в конце (например, если буфер заканчивался на \n)
        if rows and rows[-1] == '':
            rows.pop()

        # Разделяем каждую строку на ячейки (обычно по табуляции)
        parsed_data = []
        for row_text in rows:
            cells = row_text.split('\t')
            parsed_data.append(cells)

        if not parsed_data:
            logger.debug("Нет данных для вставки после разбора.")
            return

        # Получаем текущий индекс (ячейка, в которую вставляем)
        current_index = self.table_view.currentIndex()
        if not current_index.isValid():
            # Если нет текущего индекса, можно вставить, например, в A1
            current_index = self.model.index(0, 0)
            logger.debug("Текущий индекс недействителен, вставка начинается с (0, 0).")

        start_row = current_index.row()
        start_col = current_index.column()

        logger.info(f"Вставка данных из буфера в ячейку ({start_row}, {start_col}). Размер данных: {len(parsed_data)}x{len(parsed_data[0]) if parsed_data else 0}")

        # Вызываем метод модели для вставки данных
        self.model.insert_data_from_clipboard(parsed_data, start_row, start_col)

    def _on_cell_clicked(self, index: QModelIndex):
        """
        Обработчик клика по ячейке в QTableView.
        Обновляет строку формuls значением из ячейки.
        """
        if index.isValid():
            self._current_index = index
            # Получаем "сырое" значение (EditRole) для отображения в строке формул
            value = self.model.data(index, Qt.ItemDataRole.EditRole)
            self.formula_line_edit.setText(str(value) if value is not None else "")
            logger.debug(f"Ячейка ({index.row()}, {index.column()}) выбрана. Значение: {value}")

    def _on_commit_pressed(self):
        """
        Обработчик нажатия "Ввод" или кнопки "Ввод".
        Сохраняет значение из строки формул в ячейку.
        """
        if self._current_index.isValid():
            new_value = self.formula_line_edit.text()
            logger.debug(f"Сохранение значения '{new_value}' в ячейку ({self._current_index.row()}, {self._current_index.column()})")
            success = self.model.setData(self._current_index, new_value, Qt.ItemDataRole.EditRole)
            if success:
                logger.info(f"Значение ячейки ({self._current_index.row()}, {self._current_index.column()}) успешно обновлено.")
            else:
                logger.error(f"Не удалось обновить значение ячейки ({self._current_index.row()}, {self._current_index.column()}).")

    def _on_revert_pressed(self):
        """
        Обработчик нажатия кнопки "Отмена".
        Восстанавливает значение в строке формул из модели.
        """
        if self._current_index.isValid():
            current_value = self.model.data(self._current_index, Qt.ItemDataRole.EditRole)
            self.formula_line_edit.setText(str(current_value) if current_value is not None else "")
            logger.debug(f"Отмена изменений. Восстановлено значение: {current_value}")

    def _on_f2_pressed(self):
        """
        Обработчик нажатия F2.
        Переводит выбранную ячейку в режим редактирования и фокусируется на строке формул.
        """
        current_index = self.table_view.currentIndex()
        if current_index.isValid():
            self._current_index = current_index
            # Включаем режим редактирования ячейки
            self.table_view.edit(current_index)
            # Фокус на строку формул
            self.formula_line_edit.setFocus()
            # Выделяем текст для удобства
            self.formula_line_edit.selectAll()
            logger.debug(f"F2 нажат. Редактирование ячейки ({current_index.row()}, {current_index.column()})")
