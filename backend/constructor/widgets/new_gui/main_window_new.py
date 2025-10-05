"""
Новое главное окно приложения Excel Micro DB GUI.
Следует новому дизайну: таблица - центральный элемент.

Путь к файлу: backend/constructor/widgets/new_gui/main_window_new.py
"""

import logging
from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QToolBar, QStatusBar, QFileDialog,
    QMessageBox, QWidget, QHBoxLayout, QSplitter, QListWidget,
    QLineEdit, QTableView, QVBoxLayout, QHeaderView,
    QAbstractItemView, QStyleFactory, QSizePolicy
)
from PySide6.QtCore import Qt, QThread, Signal, QModelIndex, QItemSelectionModel
from PySide6.QtGui import QAction, QIcon, QKeySequence

# Импортируем AppController
from backend.core.app_controller import create_app_controller
# Импортируем logger
from backend.utils.logger import get_logger

# Импортируем вспомогательные классы (пока заглушки)
# from .table_editor_widget_new import TableEditorWidget # <-- Будет создан позже
# from .import_dialog_new import ImportDialog # <-- Будет создан позже

logger = get_logger(__name__)

# --- Временная заглушка для TableEditor ---
from PySide6.QtWidgets import QLabel
class TableEditorWidget(QWidget):
    """Временная заглушка для TableEditor."""
    def __init__(self, app_controller, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        label = QLabel("Центральная область: Табличный редактор (QTableView)\n(Пока это заглушка)")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        self.app_controller = app_controller
        
    def load_sheet(self, sheet_name: str):
        """Заглушка для загрузки листа."""
        logger.info(f"[НОВЫЙ GUI] TableEditor: Загрузка листа '{sheet_name}'")
# --- Конец заглушки ---

# --- Временная заглушка для ImportDialog ---
from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QDialogButtonBox, QPushButton
class ImportDialog(QDialog):
    """Временная заглушка для ImportDialog."""
    def __init__(self, app_controller, parent=None):
        super().__init__(parent)
        self.app_controller = app_controller
        self.setWindowTitle("Импорт данных")
        self.resize(400, 300)
        layout = QVBoxLayout(self)
        label = QLabel("Панель импорта\n(Пока это заглушка)")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        # Добавим временную кнопку "Импорт", чтобы видеть результат
        self.temp_import_btn = QPushButton("Временный Импорт (заглушка)")
        self.temp_import_btn.clicked.connect(self._temp_on_import)
        layout.addWidget(self.temp_import_btn)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
    def _temp_on_import(self):
        """Временная логика импорта."""
        logger.info("[НОВЫЙ GUI] Временная логика импорта вызвана.")
        QMessageBox.information(self, "Импорт", "Временная логика импорта (заглушка).")
# --- Конец заглушки ---


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
        
        # Атрибуты для новых компонентов
        self.project_list_widget = None
        self.table_editor_widget = None # <-- Новый атрибут
        self.formula_bar = None
        self.status_bar = None
        
        # Атрибуты для меню и действий
        self.new_project_action = None
        self.open_project_action = None
        self.save_project_action = None
        self.import_action = None
        self.export_action = None
        self.exit_action = None
        
        logger.info("Инициализация нового MainWindow...")
        
        self._setup_ui()
        self._setup_connections()
        
        # Установим начальное состояние
        self.statusBar().showMessage("Готов")

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
        self.new_project_action.setShortcut(QKeySequence.New)
        self.open_project_action = QAction("&Открыть проект", self)
        self.open_project_action.setShortcut(QKeySequence.Open)
        self.save_project_action = QAction("&Сохранить проект", self)
        self.save_project_action.setShortcut(QKeySequence.Save)
        self.exit_action = QAction("&Выход", self)
        self.exit_action.setShortcut(QKeySequence.Quit)
        
        file_menu.addAction(self.new_project_action)
        file_menu.addAction(self.open_project_action)
        file_menu.addAction(self.save_project_action)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_action)
        
        import_menu = menu_bar.addMenu("&Импорт")
        self.import_action = QAction("&Импорт данных...", self)
        self.import_action.setShortcut(QKeySequence("Ctrl+I"))
        import_menu.addAction(self.import_action)
        
        export_menu = menu_bar.addMenu("&Экспорт")
        self.export_action = QAction("&Экспорт проекта...", self)
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
        
        # --- Создание центрального виджета с разделителем ---
        central_widget = QWidget(self)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Создаем горизонтальный сплиттер
        splitter = QSplitter(Qt.Horizontal, self)
        
        # --- Создание списка проектов ---
        self.project_list_widget = QListWidget(self)
        self.project_list_widget.setFixedWidth(200)
        self.project_list_widget.addItems([
            "Проект 1", "Проект 2", "Проект 3 (пустой)"
        ])
        # ---------------------------------
        
        # --- Создание табличного редактора ---
        # self.table_editor_widget = TableEditorWidget(self.app_controller, self)
        # Временно используем заглушку
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
        self.status_bar.showMessage("Готов")
        # ---------------------------------

    def _setup_connections(self):
        """
        Подключает сигналы к слотам.
        """
        # --- Подключение сигналов меню и тулбара ---
        self.new_project_action.triggered.connect(self._on_new_project_triggered)
        self.open_project_action.triggered.connect(self._on_open_project_triggered)
        self.save_project_action.triggered.connect(self._on_save_project_triggered)
        self.import_action.triggered.connect(self._on_import_triggered)
        self.export_action.triggered.connect(self._on_export_triggered)
        self.exit_action.triggered.connect(self.close)
        # -------------------------------------------
        
        # --- Подключение сигналов списка проектов ---
        # TODO: Реализовать логику выбора проекта
        # self.project_list_widget.currentTextChanged.connect(self._on_project_selected)
        # ---------------------------------------------
        
        # --- Подключение сигналов табличного редактора ---
        # TODO: Реализовать логику выбора листа в TableEditor
        # self.table_editor_widget.sheet_selected.connect(self._on_sheet_selected)
        # ------------------------------------------------

    # --- Обработчики действий меню и тулбара ---
    def _on_new_project_triggered(self):
        """Обработчик создания нового проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Новый проект'")
        self.status_bar.showMessage("Создание нового проекта...")
        # TODO: Реализовать логику создания нового проекта
        QMessageBox.information(self, "Новый проект", "Логика создания нового проекта (пока заглушка).")
        self.status_bar.showMessage("Готов")

    def _on_open_project_triggered(self):
        """Обработчик открытия существующего проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Открыть проект'")
        self.status_bar.showMessage("Открытие проекта...")
        # TODO: Реализовать логику открытия проекта
        # project_path_str = QFileDialog.getExistingDirectory(
        #     self, "Открыть проект", "", options=QFileDialog.Option.DontUseNativeDialog
        # )
        # if project_path_str:
        #     # ... логика открытия
        QMessageBox.information(self, "Открыть проект", "Логика открытия проекта (пока заглушка).")
        self.status_bar.showMessage("Готов")

    def _on_save_project_triggered(self):
        """Обработчик сохранения проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Сохранить проект'")
        self.status_bar.showMessage("Сохранение проекта...")
        # TODO: Реализовать логику сохранения проекта
        QMessageBox.information(self, "Сохранить проект", "Логика сохранения проекта (пока заглушка).")
        self.status_bar.showMessage("Готов")

    def _on_import_triggered(self):
        """Обработчик импорта данных."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Импорт данных'")
        self.status_bar.showMessage("Открытие панели импорта...")
        
        # Создаем и показываем диалог импорта
        # import_dialog = ImportDialog(self.app_controller, self)
        # Временно используем заглушку
        import_dialog = ImportDialog(self.app_controller, self)
        
        # Показываем модально
        result = import_dialog.exec()
        if result == QDialog.Accepted:
            logger.info("[НОВЫЙ GUI] Импорт завершен успешно (по данным диалога).")
            self.status_bar.showMessage("Импорт завершен.")
            # TODO: Обновить список проектов и/или данные в TableEditor
        else:
            logger.info("[НОВЫЙ GUI] Импорт отменен пользователем.")
            self.status_bar.showMessage("Импорт отменен.")

    def _on_export_triggered(self):
        """Обработчик экспорта проекта."""
        logger.info("[НОВЫЙ GUI] Выбрано 'Экспорт проекта'")
        self.status_bar.showMessage("Экспорт проекта...")
        # TODO: Реализовать логику экспорта проекта
        QMessageBox.information(self, "Экспорт проекта", "Логика экспорта проекта (пока заглушка).")
        self.status_bar.showMessage("Готов")
    # ------------------------------------------


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
    app.setStyle(QStyleFactory.create("Fusion")) # Установим стиль для лучшего вида
    
    window = MainWindowNew()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    run_new_gui()
