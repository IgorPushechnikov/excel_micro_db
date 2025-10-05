# backend/constructor/widgets/new_gui/import_dialog_new.py
"""
Диалог импорта данных Excel в проект.
Позволяет выбрать файл и режим импорта.
"""

import logging
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QDialogButtonBox,
    QPushButton, QLineEdit, QFileDialog, QMessageBox, QCheckBox, QGroupBox
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
        file_group_box = QGroupBox(title="Файл Excel для импорта", parent=self)
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
        options_group_box = QGroupBox(title="Опции импорта", parent=self)
        options_layout = QVBoxLayout(options_group_box)

        # Чекбокс логирования
        self.logging_checkbox = QCheckBox("Включить логирование во время импорта", self)
        self.logging_checkbox.setChecked(True)
        options_layout.addWidget(self.logging_checkbox)

        # Селектор режима импорта (обновлённый)
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
        assert self.button_box is not None
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText("Импортировать")
        assert self.button_box is not None
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Отмена")
        
        main_layout.addWidget(self.button_box)
        # ------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        # Добавим assert, чтобы Pylance знал, что виджеты инициализированы
        assert self.browse_button is not None
        self.browse_button.clicked.connect(self._on_browse_clicked)
        assert self.button_box is not None
        self.button_box.accepted.connect(self._on_accept)
        assert self.button_box is not None
        self.button_box.rejected.connect(self.reject)
        
        # Подключение сигнала СЕЛЕКТОРА режима (type_selected больше нет)
        assert self.mode_selector is not None
        self.mode_selector.mode_selected.connect(self._on_import_mode_selected) # <-- Изменено: только mode_selected

    # --- Обработчики событий UI ---
    def _on_browse_clicked(self):
        """Обработчик нажатия кнопки 'Обзор...'.'"""
        # Добавим assert
        assert self.file_path_line_edit is not None

        file_path_str, ok = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "",
            "Excel Files (*.xlsx *.xls);;Все файлы (*)",
            options=QFileDialog.Option.DontUseNativeDialog
        )
        if ok and file_path_str:
            self.file_path_line_edit.setText(file_path_str)

    def _on_accept(self):
        """Обработчик нажатия кнопки 'Импортировать'.'"""
        # Добавим assert для всех виджетов, к которым обращаемся
        assert self.file_path_line_edit is not None
        file_path_str = self.file_path_line_edit.text().strip()
        if not file_path_str:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите файл Excel для импорта.")
            return

        file_path = Path(file_path_str)
        if not file_path.exists():
            QMessageBox.critical(self, "Ошибка", f"Файл не найден: {file_path}")
            return

        # Получаем выбранный РЕЖИМ (объединённый)
        assert self.mode_selector is not None
        import_mode_key = self.mode_selector.get_selected_mode() # <-- Изменено: один метод
        
        # Получаем состояние логирования
        assert self.logging_checkbox is not None
        is_logging_enabled = self.logging_checkbox.isChecked()
        
        # Получаем имя проекта
        assert self.project_name_line_edit is not None
        project_name = self.project_name_line_edit.text().strip()

        logger.info(
            f"Начало импорта: файл={file_path}, "
            f"режим={import_mode_key}, " # <-- Изменено: теперь один ключ
            f"логирование={is_logging_enabled}, проект={project_name}"
        )

        # TODO: Вызвать AppController для выполнения импорта, используя import_mode_key
        # Это может быть синхронный вызов или асинхронный через QThread/worker
        # Например:
        # success = self.app_controller.import_data_by_key(
        #     file_path, import_mode_key, 
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
            f"Режим: {import_mode_key}\n"
            f"Логирование: {'Вкл' if is_logging_enabled else 'Выкл'}\n"
            f"Проект: {project_name or 'Текущий'}"
        )
        self.accept() # Закрываем диалог как успешно завершённый

    def _on_import_mode_selected(self, import_mode_key: str): # <-- Изменено: принимает ключ режима
        """Обработчик выбора режима импорта.'"""
        logger.debug(f"Выбран режим импорта: {import_mode_key}")
        # Логика обновления UI в зависимости от режима может быть здесь
        # (например, включение/отключения определённых опций, если они снова появятся)
    # -----------------------------

    # --- Методы для получения данных из диалога ---
    # def get_import_type(self) -> str: # <-- Удалено: больше неактуально
    #     """Возвращает выбранный тип импорта.'"""
    #     # Добавим assert
    #     assert self.mode_selector is not None
    #     return self.mode_selector.get_selected_type()

    def get_import_mode_key(self) -> str: # <-- Новое: возвращает ключ объединённого режима
        """Возвращает ключ выбранного режима импорта.'"""
        # Добавим assert
        assert self.mode_selector is not None
        return self.mode_selector.get_selected_mode()

    # def get_import_mode(self) -> str: # <-- Изменено: теперь возвращает тот же ключ
    #     """Возвращает выбранный режим импорта.'"""
    #     # Логично возвращать тот же ключ, что и get_import_mode_key
    #     return self.get_import_mode_key()

    def get_file_path(self) -> Optional[Path]:
        """Возвращает путь к выбранному файлу.'"""
        # Добавим assert
        assert self.file_path_line_edit is not None
        path_str = self.file_path_line_edit.text().strip()
        return Path(path_str) if path_str else None

    def is_logging_enabled(self) -> bool:
        """Проверяет, включено ли логирование.'"""
        # Добавим assert
        assert self.logging_checkbox is not None
        return self.logging_checkbox.isChecked()

    def get_project_name(self) -> str:
        """Возвращает имя проекта.'"""
        # Добавим assert
        assert self.project_name_line_edit is not None
        return self.project_name_line_edit.text().strip()
    # ------------------------------------------------