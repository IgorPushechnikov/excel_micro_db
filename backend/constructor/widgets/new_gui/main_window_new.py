# backend/constructor/widgets/new_gui/main_window_new.py
"""
Новое главное окно приложения Excel Micro DB GUI.
Следует новому дизайну: таблица - центральный элемент.

Путь к файлу: backend/constructor/widgets/new_gui/main_window_new.py
"""

import logging
from pathlib import Path
from typing import Optional
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QToolBar, QStatusBar, QFileDialog,
    QMessageBox, QWidget, QHBoxLayout, QSplitter, QListWidget,
    QLineEdit, QTableView, QVBoxLayout, QHeaderView,
    QAbstractItemView, QStyleFactory, QSizePolicy, QDialog, QDialogButtonBox, QPushButton, QLabel, QGroupBox, QFormLayout, QGridLayout, QCheckBox
)
from PySide6.QtCore import Qt, QThread, Signal, QModelIndex, QItemSelectionModel
from PySide6.QtGui import QAction, QIcon, QKeySequence, QStandardItemModel, QStandardItem

# Импортируем AppController
from backend.core.app_controller import create_app_controller
# Импортируем logger
from backend.utils.logger import get_logger

# Импортируем вспомогательные классы
# Заменяем заглушки на реальные импорты
from .table_editor_widget_new import TableEditorWidget
from .import_dialog_new import ImportDialog
# from .import_mode_selector_new import ImportModeSelector # <-- Не нужен напрямую, используется внутри ImportDialog

logger = get_logger(__name__)

# --- УДАЛЕНО: Временные заглушки для TableEditor и ImportDialog ---
# class TableEditorWidget(QWidget): ...
# class ImportDialog(QDialog): ...
# --- КОНЕЦ УДАЛЕНИЯ ---


class MainWindowNew(QMainWindow):
    """
    Новое главное окно GUI приложения.
    """

    # Сигнал для уведомления о выборе листа
    sheet_selected = Signal(str)

    def __init__(self):
        """
        Инициализирует новое главное окно.
        """
        super().__init__()
        self.app_controller = create_app_controller()
        self.current_project_path = None
        self.current_sheet_name = None

        # --- Атрибуты виджета ---
        self.project_list_widget: Optional[QListWidget] = None
        self.table_editor_widget: Optional[TableEditorWidget] = None
        self.formula_bar: Optional[QLineEdit] = None
        self.status_bar: Optional[QStatusBar] = None

        # Атрибуты для меню и действий
        self.new_project_action: Optional[QAction] = None
        self.open_project_action: Optional[QAction] = None
        self.save_project_action: Optional[QAction] = None
        self.import_action: Optional[QAction] = None
        self.export_action: Optional[QAction] = None
        self.exit_action: Optional[QAction] = None

        # Атрибуты для импорта/экспорта
        self.import_worker: Optional[ImportWorker] = None # <-- Если ImportWorker будет создан
        self.export_worker: Optional[ExportWorker] = None # <-- Если ExportWorker будет создан
        # --- КОНЕЦ АТРИБУТОВ ---

        logger.info("Инициализация нового MainWindow...")

        self._setup_ui()
        self._setup_connections()

        # Установим начальное состояние
        self.statusBar().showMessage("Готов.")

    def _setup_ui(self):
        """
        Создаёт элементы интерфейса нового дизайна.
        """
        self.setWindowTitle("Excel Micro DB GUI - Новый дизайн")
        self.setGeometry(100, 100, 1200, 800)

        # --- Создание меню ---
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("&Файл")
        self.new_project_action = QAction("&Новый проект", self)
        # ИСПРАВЛЕНО: QKeySequence.New -> QKeySequence.StandardKey.New
        self.new_project_action.setShortcut(QKeySequence(QKeySequence.StandardKey.New))
        self.open_project_action = QAction("&Открыть проект", self)
        # ИСПРАВЛЕНО: QKeySequence.Open -> QKeySequence.StandardKey.Open
        self.open_project_action.setShortcut(QKeySequence(QKeySequence.StandardKey.Open))
        self.save_project_action = QAction("&Сохранить проект", self)
        # ИСПРАВЛЕНО: QKeySequence.Save -> QKeySequence.StandardKey.Save
        self.save_project_action.setShortcut(QKeySequence(QKeySequence.StandardKey.Save))
        self.exit_action = QAction("&Выход", self)
        # ИСПРАВЛЕНО: QKeySequence.Quit -> QKeySequence.StandardKey.Quit
        self.exit_action.setShortcut(QKeySequence(QKeySequence.StandardKey.Quit))

        file_menu.addAction(self.new_project_action)
        file_menu.addAction(self.open_project_action)
        file_menu.addAction(self.save_project_action)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_action)

        import_menu = menu_bar.addMenu("&Импорт")
        self.import_action = QAction("&Импорт данных...", self)
        # ИСПРАВЛЕНО: QKeySequence("Ctrl+I") -> QKeySequence.StandardKey
        self.import_action.setShortcut(QKeySequence("Ctrl+I"))
        import_menu.addAction(self.import_action)

        export_menu = menu_bar.addMenu("&Экспорт")
        self.export_action = QAction("&Экспорт проекта...", self)
        # ИСПРАВЛЕНО: QKeySequence("Ctrl+E") -> QKeySequence.StandardKey
        self.export_action.setShortcut(QKeySequence("Ctrl+E"))
        export_menu.addAction(self.export_action)
        # ----------------------

        # --- Создание панели инструментов ---
        tool_bar = self.addToolBar("Основная")
        tool_bar.addAction(self.new_project_action)
        tool_bar.addAction(self.open_project_action)
        tool_bar.addAction(self.save_project_action)
        tool_bar.addSeparator()
        tool_bar.addAction(self.import_action)
        tool_bar.addAction(self.export_action)
        # -----------------------------------

        # --- Создание центрального виджета с разделителем (Splitter) ---
        central_widget = QWidget(self)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Создаем горизонтальный сплиттер
        splitter = QSplitter(Qt.Orientation.Horizontal, self) # <-- ИСПРАВЛЕНО: Qt.Horizontal -> Qt.Orientation.Horizontal

        # --- Создание списка проектов ---
        self.project_list_widget = QListWidget(self)
        self.project_list_widget.setFixedWidth(200)
        self.project_list_widget.addItems([
            "Проект 1", "Проект 2", "Проект 3 (пустой)"
        ])
        # ---------------------------------

        # --- Создание табличного редактора ---
        # Заменяем заглушку на реальный TableEditorWidget
        self.table_editor_widget = TableEditorWidget(self.app_controller, self)
        # -------------------------------------

        # Добавляем виджеты в сплиттер
        splitter.addWidget(self.project_list_widget)
        splitter.addWidget(self.table_editor_widget)
        splitter.setSizes([200, 1000]) # Примерное соотношение ширины

        main_layout.addWidget(splitter)
        self.setCentralWidget(central_widget)
        # ---------------------------------------

        # --- Создание строки состояния ---
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов.")
        # ---------------------------------

    def _setup_connections(self):
        """
        Подключает сигналы к слотам.
        """
        # --- Подключение сигналов меню и тулбара ---
        if self.new_project_action:
            self.new_project_action.triggered.connect(self._on_new_project_triggered)
        if self.open_project_action:
            self.open_project_action.triggered.connect(self._on_open_project_triggered)
        if self.save_project_action:
            self.save_project_action.triggered.connect(self._on_save_project_triggered)
        if self.import_action:
            self.import_action.triggered.connect(self._on_import_triggered)
        if self.export_action:
            self.export_action.triggered.connect(self._on_export_triggered)
        if self.exit_action:
            self.exit_action.triggered.connect(self.close)
        # -------------------------------------------

        # --- Подключение сигналов списка проектов ---
        # TODO: Реализовать логику выбора проекта
        # if self.project_list_widget:
        #     self.project_list_widget.currentTextChanged.connect(self._on_project_selected)
        # ---------------------------------------------

        # --- Подключение сигналов табличного редактора ---
        # TODO: Реализовать логику выбора листа в TableEditor
        # if self.table_editor_widget:
        #     self.table_editor_widget.sheet_selected.connect(self._on_sheet_selected)
        # ------------------------------------------------

    # --- Обработчики действий меню и тулбара ---
    def _on_new_project_triggered(self):
        """Обработчик создания нового проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Новый проект'")
        self.statusBar().showMessage("Создание нового проекта...")
        # TODO: Реализовать логику создания нового проекта
        QMessageBox.information(self, "Новый проект", "Логика создания нового проекта (пока заглушка).")
        self.statusBar().showMessage("Готов.")

    def _on_open_project_triggered(self):
        """Обработчик открытия существующего проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Открыть проект'")
        self.statusBar().showMessage("Открытие проекта...")
        # TODO: Реализовать логику открытия проекта
        # project_path_str = QFileDialog.getExistingDirectory(
        #     self, "Открыть проект", "", options=QFileDialog.Option.DontUseNativeDialog
        # )
        # if project_path_str:
        #     # ... логика открытия
        QMessageBox.information(self, "Открыть проект", "Логика открытия проекта (пока заглушка).")
        self.statusBar().showMessage("Готов.")

    def _on_save_project_triggered(self):
        """Обработчик сохранения проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Сохранить проект'")
        self.statusBar().showMessage("Сохранение проекта...")
        # TODO: Реализовать логику сохранения проекта
        QMessageBox.information(self, "Сохранить проект", "Логика сохранения проекта (пока заглушка).")
        self.statusBar().showMessage("Готов.")

    def _on_import_triggered(self):
        """Обработчик импорта данных."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Импорт данных...'")
        self.statusBar().showMessage("Открытие панели импорта...")

        # Создаем и показываем диалог импорта
        # Заменяем заглушку на реальный ImportDialog
        import_dialog = ImportDialog(self.app_controller, self)

        # Показываем модально
        # ИСПРАВЛЕНО: QDialog.Accepted -> QDialog.DialogCode.Accepted
        result = import_dialog.exec()
        if result == QDialog.DialogCode.Accepted: # <-- ИСПРАВЛЕНО
            logger.info("[НОВЫЙ GUI] Импорт подтверждён в диалоге.")

            # Получаем данные из диалога
            file_path = import_dialog.get_file_path()
            import_type = import_dialog.get_import_type()
            import_mode = import_dialog.get_import_mode()
            is_logging_enabled = import_dialog.is_logging_enabled()
            project_name = import_dialog.get_project_name() # Можно использовать позже

            if not file_path or not file_path.exists():
                logger.error(f"[НОВЫЙ GUI] Неверный путь к файлу импорта: {file_path}")
                QMessageBox.critical(self, "Ошибка", f"Неверный путь к файлу импорта: {file_path}")
                return

            # Сохраняем выбранные тип и режим для будущего использования
            # (если нужно сохранять между сессиями или для других целей)
            # self.selected_import_type = import_type
            # self.selected_import_mode = import_mode

            # Устанавливаем состояние логирования через AppController
            self.app_controller.set_logging_enabled(is_logging_enabled)
            logger.info(f"[НОВЫЙ GUI] Логирование {'включено' if is_logging_enabled else 'отключено'} для импорта.")

            # Создаем рабочий поток для импорта (если ImportWorker будет создан)
            # self.import_worker = ImportWorker(self.app_controller, str(file_path), import_type, import_mode)
            # self.import_worker.finished.connect(self._on_import_finished)
            # self.import_worker.progress.connect(self._on_import_progress) # <-- НОВОЕ: Подключаем progress

            # self.import_worker.start()
            self.statusBar().showMessage(f"Начат импорт {file_path.name}...")
            logger.info(f"[НОВЫЙ GUI] Начат импорт: файл={file_path}, тип={import_type}, режим={import_mode}")

            # Пока что просто покажем сообщение
            QMessageBox.information(
                self, "Импорт",
                f"Импорт запущен (заглушка).\n"
                f"Файл: {file_path.name}\n"
                f"Тип: {import_type}\n"
                f"Режим: {import_mode}\n"
                f"Логирование: {'Вкл' if is_logging_enabled else 'Выкл'}\n"
                f"Проект: {project_name or 'Текущий'}"
            )

        else:
            logger.info("[НОВЫЙ GUI] Импорт отменён пользователем.")
            self.statusBar().showMessage("Импорт отменён.")

    def _on_export_triggered(self):
        """Обработчик экспорта проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Экспорт проекта'")
        self.statusBar().showMessage("Экспорт проекта...")
        # TODO: Реализовать логику экспорта проекта
        QMessageBox.information(self, "Экспорт проекта", "Логика экспорта проекта (пока заглушка).")
        self.statusBar().showMessage("Готов.")
    # ------------------------------------------

    # --- Обработчики прогресса импорта/экспорта (если поддерживается) ---
    def _on_import_progress(self, value, message):
        """Обработчик прогресса импорта (если поддерживается)."""
        # Обновляем сообщение в строке состояния
        self.statusBar().showMessage(message)

    def _on_export_progress(self, value, message):
        """Обработчик прогресса экспорта (если поддерживается)."""
        # Обновляем сообщение в строке состояния
        self.statusBar().showMessage(message)

    def _on_import_finished(self, success, message):
        """Обработчик завершения импорта."""
        # Убираем ссылку на worker, чтобы он мог быть уничтожен
        if self.import_worker:
            self.import_worker.wait() # Убедиться, что поток завершён
            self.import_worker = None

        if success:
            logger.info(f"[НОВЫЙ GUI] Импорт успешно завершён: {message}")
            self.statusBar().showMessage(f"Импорт завершён: {message}")
            # Обновляем список листов, так как могли появиться новые
            # self._update_sheet_list()
        else:
            logger.error(f"[НОВЫЙ GUI] Импорт завершился с ошибкой: {message}")
            self.statusBar().showMessage(f"Ошибка импорта: {message}")
            # QMessageBox.critical(self, "Ошибка", message) # <-- Опционально, можно оставить

    def _on_export_finished(self, success, message):
        """Обработчик завершения экспорта."""
        # Убираем ссылку на worker, чтобы он мог быть уничтожен
        if self.export_worker:
            self.export_worker.wait() # Убедиться, что поток завершён
            self.export_worker = None

        if success:
            logger.info(f"[НОВЫЙ GUI] Экспорт успешно завершён: {message}")
            self.statusBar().showMessage(f"Экспорт завершён: {message}")
        else:
            logger.error(f"[НОВЫЙ GUI] Экспорт завершился с ошибкой: {message}")
            self.statusBar().showMessage(f"Ошибка экспорта: {message}")
            # QMessageBox.critical(self, "Ошибка", message) # <-- Опционально, можно оставить
    # ---------------------------------------------------------------


# --- Функция для запуска нового GUI ---
def run_new_gui():
    """
    Функция для запуска нового GUI из скрипта.
    """
    import sys
    from PySide6.QtWidgets import QApplication

    # Настройка логирования (если нужно)
    # from backend.utils.logger import setup_logger
    # setup_logger()

    app = QApplication(sys.argv)
    # ИСПРАВЛЕНО: QStyleFactory.create("Fusion") -> без изменений, но убедимся, что стиль доступен
    app.setStyle(QStyleFactory.create("Fusion")) # Установим стиль для лучшего вида

    window = MainWindowNew()
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    run_new_gui()
