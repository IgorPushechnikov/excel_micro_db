# src/constructor/widgets/main_window.py
"""
Модуль для главного окна графического интерфейса.
"""

import logging
from pathlib import Path
from typing import Optional

# from PySide6.QtCore import Qt, Slot, Signal # Старый импорт Qt
from PySide6.QtCore import Slot, Signal
from PySide6.QtGui import QAction, QIcon
# Импортируем QtCore и QtGui напрямую для доступа к атрибутам Qt
from PySide6 import QtCore, QtGui
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QMenu, QToolBar, QStatusBar,
    QDockWidget, QTreeWidget, QTreeWidgetItem,
    QTabWidget, QWidget, QLabel, QMessageBox,
    QVBoxLayout, QHBoxLayout, QSplitter
)

# Импортируем AppController
from src.core.app_controller import AppController

# Импортируем виджеты
from src.constructor.widgets.project_explorer import ProjectExplorer
# from src.constructor.widgets.sheet_editor import SheetEditor # Пока не импортируем

# Получаем логгер
logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """
    Главное окно приложения Excel Micro DB.
    """

    def __init__(self, app_controller: AppController):
        """
        Инициализирует главное окно.

        Args:
            app_controller (AppController): Экземпляр основного контроллера приложения.
        """
        super().__init__()
        self.app_controller: AppController = app_controller
        self.project_explorer: Optional[ProjectExplorer] = None
        # self.sheet_editor: Optional[SheetEditor] = None # Пока не создаём

        self._setup_ui()
        self._create_actions()
        self._create_menus()
        self._create_toolbars()
        self._create_status_bar()
        self._connect_signals()

        self.setWindowTitle("Excel Micro DB")
        self.resize(1200, 800)
        logger.debug("MainWindow инициализировано.")

    def _setup_ui(self):
        """Настраивает пользовательский интерфейс."""
        logger.debug("Настройка UI главного окна...")

        # --- Центральный виджет ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(0, 0, 0, 0)

        # --- Project Explorer (Dock Widget) ---
        self.project_explorer_dock = QDockWidget("Проект", self)
        self.project_explorer_dock.setAllowedAreas(
            QtCore.Qt.LeftDockWidgetArea | QtCore.Qt.RightDockWidgetArea
        )
        self.project_explorer = ProjectExplorer(self.app_controller)
        self.project_explorer_dock.setWidget(self.project_explorer)
        self.addDockWidget(QtCore.Qt.LeftDockWidgetArea, self.project_explorer_dock)

        # --- Splitter для центральной области ---
        self.central_splitter = QSplitter(QtCore.Qt.Horizontal)
        central_layout.addWidget(self.central_splitter)

        # --- Tab Widget для SheetEditor ---
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        # Добавляем tab_widget в splitter
        self.central_splitter.addWidget(self.tab_widget)
        
        # --- Панель свойств (справа, например) ---
        self.properties_dock = QDockWidget("Свойства", self)
        properties_widget = QLabel("Панель свойств\n(Пока недоступна)")
        properties_widget.setAlignment(QtCore.Qt.AlignCenter)
        self.properties_dock.setWidget(properties_widget)
        self.addDockWidget(QtCore.Qt.RightDockWidgetArea, self.properties_dock)

        logger.debug("UI главного окна настроено.")

    def _create_actions(self):
        """Создаёт действия (QAction) для меню и панелей инструментов."""
        logger.debug("Создание действий...")
        # --- Файл ---
        self.action_new_project = QAction("&Новый проект", self)
        self.action_new_project.setShortcut("Ctrl+N")
        # self.action_new_project.triggered.connect(self._on_new_project)

        self.action_open_project = QAction("&Открыть проект", self)
        self.action_open_project.setShortcut("Ctrl+O")
        # self.action_open_project.triggered.connect(self._on_open_project)

        self.action_save_project = QAction("&Сохранить проект", self)
        self.action_save_project.setShortcut("Ctrl+S")
        # self.action_save_project.setEnabled(False) 

        self.action_close_project = QAction("&Закрыть проект", self)
        # self.action_close_project.triggered.connect(self._on_close_project)
        # self.action_close_project.setEnabled(False) 

        self.action_exit = QAction("&Выход", self)
        self.action_exit.setShortcut("Ctrl+Q")
        self.action_exit.triggered.connect(self.close)

        # --- Инструменты ---
        self.action_analyze_excel = QAction("&Анализировать Excel-файл", self)
        # self.action_analyze_excel.triggered.connect(self._on_analyze_excel)
        # self.action_analyze_excel.setEnabled(False) 

        self.action_export_to_excel = QAction("&Экспортировать в Excel", self)
        # self.action_export_to_excel.triggered.connect(self._on_export_to_excel)
        # self.action_export_to_excel.setEnabled(False) 

        # --- Помощь ---
        self.action_about = QAction("&О программе", self)
        # self.action_about.triggered.connect(self._on_about)

        logger.debug("Действия созданы.")

    def _create_menus(self):
        """Создаёт меню."""
        logger.debug("Создание меню...")
        menubar: QMenuBar = self.menuBar()

        # --- Меню Файл ---
        file_menu: QMenu = menubar.addMenu("&Файл")
        file_menu.addAction(self.action_new_project)
        file_menu.addAction(self.action_open_project)
        file_menu.addAction(self.action_save_project)
        file_menu.addSeparator()
        file_menu.addAction(self.action_close_project)
        file_menu.addAction(self.action_exit)

        # --- Меню Инструменты ---
        tools_menu: QMenu = menubar.addMenu("&Инструменты")
        tools_menu.addAction(self.action_analyze_excel)
        tools_menu.addAction(self.action_export_to_excel)

        # --- Меню Помощь ---
        help_menu: QMenu = menubar.addMenu("&Помощь")
        help_menu.addAction(self.action_about)

        logger.debug("Меню созданы.")

    def _create_toolbars(self):
        """Создаёт панели инструментов."""
        logger.debug("Создание панелей инструментов...")
        toolbar: QToolBar = self.addToolBar("Основные")
        toolbar.setObjectName("main_toolbar") 
        toolbar.addAction(self.action_new_project)
        toolbar.addAction(self.action_open_project)
        toolbar.addAction(self.action_analyze_excel)
        toolbar.addAction(self.action_export_to_excel)
        logger.debug("Панели инструментов созданы.")

    def _create_status_bar(self):
        """Создаёт строку состояния."""
        logger.debug("Создание строки состояния...")
        self.statusBar().showMessage("Готов")
        logger.debug("Строка состояния создана.")

    def _connect_signals(self):
        """Подключает сигналы и слоты."""
        logger.debug("Подключение сигналов...")
        # Пример подключения сигнала из ProjectExplorer
        # self.project_explorer.sheet_selected.connect(self._on_sheet_selected)
        # ... другие подключения
        logger.debug("Сигналы подключены.")

    # --- Слоты для действий ---
    # @Slot()
    # def _on_new_project(self):
    #     logger.debug("Выбрано действие: Новый проект")
    #     # TODO: Реализовать

    # @Slot()
    # def _on_open_project(self):
    #     logger.debug("Выбрано действие: Открыть проект")
    #     # TODO: Реализовать

    # @Slot()
    # def _on_analyze_excel(self):
    #     logger.debug("Выбрано действие: Анализировать Excel-файл")
    #     # TODO: Реализовать

    # @Slot()
    # def _on_export_to_excel(self):
    #     logger.debug("Выбрано действие: Экспортировать в Excel")
    #     # TODO: Реализовать

    # @Slot(str)
    # def _on_sheet_selected(self, sheet_name: str):
    #     """Слот для обработки выбора листа в ProjectExplorer."""
    #     logger.debug(f"MainWindow получил сигнал о выборе листа: {sheet_name}")
    #     # TODO: Открыть/обновить вкладку SheetEditor
    #     # Например:
    #     # if self.sheet_editor is None:
    #     #     self.sheet_editor = SheetEditor(self.app_controller)
    #     #     self.tab_widget.addTab(self.sheet_editor, sheet_name)
    #     # self.sheet_editor.load_sheet(...) # Передать данные
