# backend/constructor/widgets/new_gui/selective_import_options_dialog.py
"""
Диалог для выбора опций при выборочном импорте данных.
Позволяет выбрать листы Excel для импорта.
"""

import logging
from typing import List, Optional
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QDialogButtonBox,
    QPushButton, QLineEdit, QFileDialog, QMessageBox, QCheckBox, QGroupBox,
    QListWidget, QListWidgetItem
)
from PySide6.QtCore import Qt, Signal

from backend.utils.logger import get_logger

logger = get_logger(__name__)

class SelectiveImportOptionsDialog(QDialog):
    """
    Диалог выбора опций для выборочного импорта.
    """

    def __init__(self, available_sheet_names: List[str], parent=None):
        """
        Инициализирует диалог.

        Args:
            available_sheet_names (List[str]): Список имён доступных листов в файле.
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self.available_sheet_names = available_sheet_names
        self.selected_sheet_names: List[str] = []

        self.setWindowTitle("Выбор опций для выборочного импорта")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.resize(400, 300)

        # --- Атрибуты диалога ---
        self.sheets_list_widget: Optional[QListWidget] = None
        self.select_all_button: Optional[QPushButton] = None
        self.deselect_all_button: Optional[QPushButton] = None
        self.button_box: Optional[QDialogButtonBox] = None
        # -----------------------

        self._setup_ui()
        self._setup_connections()
        self._populate_sheets()

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        main_layout = QVBoxLayout(self)

        # --- Группа выбора листов ---
        sheets_group_box = QGroupBox(title="Листы Excel для импорта", parent=self)
        sheets_layout = QVBoxLayout(sheets_group_box)

        # Кнопки "Выбрать все" / "Снять выделение"
        buttons_layout = QHBoxLayout()
        self.select_all_button = QPushButton("Выбрать все", self)
        self.deselect_all_button = QPushButton("Снять выделение", self)
        buttons_layout.addWidget(self.select_all_button)
        buttons_layout.addWidget(self.deselect_all_button)
        buttons_layout.addStretch()
        sheets_layout.addLayout(buttons_layout)

        # Список листов с чекбоксами
        self.sheets_list_widget = QListWidget(self)
        sheets_layout.addWidget(self.sheets_list_widget)

        main_layout.addWidget(sheets_group_box)
        # --------------------------

        # --- Кнопки OK/Cancel ---
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self
        )
        # Установим текст кнопок на русский
        assert self.button_box is not None
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText("Импортировать")
        assert self.button_box is not None
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Отмена")
        
        main_layout.addWidget(self.button_box)
        # ------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        # --- Подключение сигналов кнопок ---
        assert self.select_all_button is not None
        self.select_all_button.clicked.connect(self._on_select_all_clicked)
        assert self.deselect_all_button is not None
        self.deselect_all_button.clicked.connect(self._on_deselect_all_clicked)
        assert self.button_box is not None
        self.button_box.accepted.connect(self._on_accept)
        assert self.button_box is not None
        self.button_box.rejected.connect(self.reject)
        # ----------------------------------

    def _populate_sheets(self):
        """Заполняет список листов."""
        assert self.sheets_list_widget is not None
        self.sheets_list_widget.clear()
        for sheet_name in self.available_sheet_names:
            item = QListWidgetItem(sheet_name)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.sheets_list_widget.addItem(item)

    def _on_select_all_clicked(self):
        """Обработчик нажатия кнопки "Выбрать все"."""
        assert self.sheets_list_widget is not None
        for i in range(self.sheets_list_widget.count()):
            item = self.sheets_list_widget.item(i)
            item.setCheckState(Qt.CheckState.Checked)

    def _on_deselect_all_clicked(self):
        """Обработчик нажатия кнопки "Снять выделение"."""
        assert self.sheets_list_widget is not None
        for i in range(self.sheets_list_widget.count()):
            item = self.sheets_list_widget.item(i)
            item.setCheckState(Qt.CheckState.Unchecked)

    def _on_accept(self):
        """Обработчик нажатия кнопки "Импортировать"."""
        assert self.sheets_list_widget is not None
        self.selected_sheet_names = []
        for i in range(self.sheets_list_widget.count()):
            item = self.sheets_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                self.selected_sheet_names.append(item.text())
        
        if not self.selected_sheet_names:
            QMessageBox.warning(self, "Предупреждение", "Пожалуйста, выберите хотя бы один лист для импорта.")
            return

        logger.debug(f"Выбраны листы для выборочного импорта: {self.selected_sheet_names}")
        self.accept() # Закрываем диалог как успешно завершённый

    def get_selected_sheet_names(self) -> List[str]:
        """
        Возвращает список выбранных имён листов.

        Returns:
            List[str]: Список имён листов.
        """
        return self.selected_sheet_names

# --- Пример использования (для тестирования) ---
# if __name__ == "__main__":
#     import sys
#     from PySide6.QtWidgets import QApplication
#
#     app = QApplication(sys.argv)
#
#     # Пример данных
#     sheet_names = ["Лист1", "Лист2", "Данные", "Сводка"]
#
#     dialog = SelectiveImportOptionsDialog(sheet_names)
#     if dialog.exec() == QDialog.DialogCode.Accepted:
#         selected = dialog.get_selected_sheet_names()
#         print(f"Выбраны листы: {selected}")
#     else:
#         print("Диалог отменён.")
#
#     sys.exit(app.exec())
