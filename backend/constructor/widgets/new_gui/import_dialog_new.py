# backend/constructor/widgets/new_gui/import_dialog_new.py
"""
Диалог импорта данных Excel в проект.
Позволяет выбрать файл, тип и режим импорта.
"""

import logging
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGridLayout,
    QPushButton, QLineEdit, QFileDialog, QMessageBox, QCheckBox,
    QDialogButtonBox, QGroupBox, QLabel
)
from PySide6.QtCore import Qt, Signal

# Импортируем AppController
from backend.core.app_controller import create_app_controller
from backend.utils.logger import get_logger

# Импортируем селектор режима импорта
from .import_mode_selector_new import ImportModeSelector

logger = get_logger(__name__)

class ImportDialog(QDialog):
    """
    Диалог импорта данных Excel.
    """

    def __init__(self, app_controller, parent=None):
        """
        Инициализирует диалог импорта.

        Args:
            app_controller: Экземпляр AppController.
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self.app_controller = app_controller
        self.setWindowTitle("Импорт данных Excel")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.resize(500, 400)

        # --- Атрибуты диалога ---
        self.file_path_line_edit: Optional[QLineEdit] = None
        self.browse_button: Optional[QPushButton] = None
        self.logging_checkbox: Optional[QCheckBox] = None
        self.mode_selector: Optional[ImportModeSelector] = None
        self.project_name_line_edit: Optional[QLineEdit] = None
        self.button_box: Optional[QDialogButtonBox] = None
        # -----------------------

        self._setup_ui()
        self._setup_connections()

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        main_layout = QVBoxLayout(self)

        # --- Форма выбора файла ---
        file_group_box = QGroupBox("Файл Excel для импорта")
        file_layout = QFormLayout(file_group_box)

        self.file_path_line_edit = QLineEdit(self)
        self.file_path_line_edit.setPlaceholderText("Путь к .xlsx или .xls файлу...")
        self.browse_button = QPushButton("Обзор...", self)
        self.browse_button.setFixedWidth(100)

        file_browse_layout = QHBoxLayout()
        file_browse_layout.addWidget(self.file_path_line_edit)
        file_browse_layout.addWidget(self.browse_button)

        file_layout.addRow("Файл:", file_browse_layout)
        main_layout.addWidget(file_group_box)
        # --------------------------

        # --- Опции импорта ---
        options_group_box = QGroupBox("Опции импорта")
        options_layout = QVBoxLayout(options_group_box)

        # Чекбокс логирования
        self.logging_checkbox = QCheckBox("Включить логирование во время импорта", self)
        self.logging_checkbox.setChecked(True)
        options_layout.addWidget(self.logging_checkbox)

        # Селектор типа и режима импорта
        self.mode_selector = ImportModeSelector(self)
        options_layout.addWidget(self.mode_selector)

        # Имя проекта (если нужно создать новый)
        project_layout = QFormLayout()
        self.project_name_line_edit = QLineEdit(self)
        self.project_name_line_edit.setPlaceholderText("Имя проекта (если создается новый)...")
        project_layout.addRow("Проект:", self.project_name_line_edit)
        options_layout.addLayout(project_layout)

        main_layout.addWidget(options_group_box)
        # --------------------

        # --- Кнопки OK/Cancel ---
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self
        )
        # Установим текст кнопок на русский
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText("Импортировать")
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Отмена")
        
        main_layout.addWidget(self.button_box)
        # ------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        self.browse_button.clicked.connect(self._on_browse_clicked)
        self.button_box.accepted.connect(self._on_accept)
        self.button_box.rejected.connect(self.reject)
        
        # Подключение сигналов селектора режима
        self.mode_selector.type_selected.connect(self._on_import_type_selected)
        self.mode_selector.mode_selected.connect(self._on_import_mode_selected)

    # --- Обработчики событий UI ---
    def _on_browse_clicked(self):
        """Обработчик нажатия кнопки 'Обзор...'."""
        file_path_str, ok = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "",
            "Excel Files (*.xlsx *.xls);;Все файлы (*)",
            options=QFileDialog.Option.DontUseNativeDialog
        )
        if ok and file_path_str:
            self.file_path_line_edit.setText(file_path_str)

    def _on_accept(self):
        """Обработчик нажатия кнопки 'Импортировать'."""
        file_path_str = self.file_path_line_edit.text().strip()
        if not file_path_str:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите файл Excel для импорта.")
            return

        file_path = Path(file_path_str)
        if not file_path.exists():
            QMessageBox.critical(self, "Ошибка", f"Файл не найден: {file_path}")
            return

        # Получаем выбранный тип и режим
        import_type = self.mode_selector.get_selected_type()
        import_mode = self.mode_selector.get_selected_mode()
        
        # Получаем состояние логирования
        is_logging_enabled = self.logging_checkbox.isChecked()
        
        # Получаем имя проекта
        project_name = self.project_name_line_edit.text().strip()

        logger.info(
            f"Начало импорта: файл={file_path}, тип={import_type}, "
            f"режим={import_mode}, логирование={is_logging_enabled}, проект={project_name}"
        )

        # TODO: Вызвать AppController для выполнения импорта
        # Это может быть синхронный вызов или асинхронный через QThread/worker
        # Например:
        # success = self.app_controller.import_data(
        #     file_path, import_type, import_mode, 
        #     enable_logging=is_logging_enabled, project_name=project_name
        # )
        # if success:
        #     self.accept()
        # else:
        #     QMessageBox.critical(self, "Ошибка", "Импорт не удался.")
        
        # Пока что просто покажем сообщение и закроем диалог
        QMessageBox.information(
            self, "Импорт", 
            f"Импорт запущен (заглушка).\n"
            f"Файл: {file_path.name}\n"
            f"Тип: {import_type}\n"
            f"Режим: {import_mode}\n"
            f"Логирование: {'Вкл' if is_logging_enabled else 'Выкл'}\n"
            f"Проект: {project_name or 'Текущий'}"
        )
        self.accept() # Закрываем диалог как успешно завершённый

    def _on_import_type_selected(self, import_type: str):
        """Обработчик выбора типа импорта."""
        logger.debug(f"Выбран тип импорта: {import_type}")
        # Логика обновления UI в зависимости от типа может быть здесь
        # Например, включение/отключение определённых опций

    def _on_import_mode_selected(self, import_mode: str):
        """Обработчик выбора режима импорта."""
        logger.debug(f"Выбран режим импорта: {import_mode}")
        # Логика обновления UI в зависимости от режима может быть здесь
    # -----------------------------

    # --- Методы для получения данных из диалога ---
    def get_file_path(self) -> Optional[Path]:
        """Возвращает путь к выбранному файлу."""
        path_str = self.file_path_line_edit.text().strip()
        return Path(path_str) if path_str else None

    def is_logging_enabled(self) -> bool:
        """Проверяет, включено ли логирование."""
        return self.logging_checkbox.isChecked() if self.logging_checkbox else True

    def get_import_type(self) -> str:
        """Возвращает выбранный тип импорта."""
        return self.mode_selector.get_selected_type() if self.mode_selector else "all_data"

    def get_import_mode(self) -> str:
        """Возвращает выбранный режим импорта."""
        return self.mode_selector.get_selected_mode() if self.mode_selector else "all"

    def get_project_name(self) -> str:
        """Возвращает имя проекта."""
        return self.project_name_line_edit.text().strip() if self.project_name_line_edit else ""
    # ------------------------------------------------
