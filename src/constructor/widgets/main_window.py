# src/constructor/widgets/main_window.py
"""
Модуль для главного окна графического интерфейса с использованием PySide6-Fluent-Widgets.
"""

import logging
from pathlib import Path
from typing import Optional

# --- Импорт из PySide6 ---
from PySide6.QtCore import Slot, Signal
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QMessageBox
)
# --- КОНЕЦ Импорта из PySide6 ---

# --- НОВОЕ: Импорт из qfluentwidgets ---
from qfluentwidgets import FluentWindow, SubtitleLabel, setFont
# --- КОНЕЦ НОВОГО ---

# Импортируем AppController
from src.core.app_controller import AppController

# Импортируем виджеты
# from src.constructor.widgets.project_explorer import ProjectExplorer # Пока не используем напрямую, FluentWindow управляет навигацией
# from src.constructor.widgets.sheet_editor import SheetEditor # Пока не импортируем

# Получаем логгер
logger = logging.getLogger(__name__)


class MainWindow(FluentWindow):
    """
    Главное окно приложения Excel Micro DB с Fluent дизайном.
    Наследуется от FluentWindow для использования Fluent Navigation Interface.
    """

    def __init__(self, app_controller: AppController):
        """
        Инициализирует главное окно с Fluent дизайном.

        Args:
            app_controller (AppController): Экземпляр основного контроллера приложения.
        """
        super().__init__()
        self.app_controller: AppController = app_controller
        # self.project_explorer: Optional[ProjectExplorer] = None # FluentWindow управляет
        # self.sheet_editor: Optional[SheetEditor] = None # FluentWindow управляет

        self._setup_ui()
        # self._create_actions() # FluentWindow использует свои собственные действия
        # self._create_menus() # FluentWindow использует NavigationInterface
        # self._create_toolbars() # FluentWindow использует свои собственные панели
        # self._create_status_bar() # FluentWindow использует InfoBar
        # self._connect_signals() # Подключение сигналов остаётся

        self.setWindowTitle("Excel Micro DB")
        self.resize(1200, 800)
        logger.debug("MainWindow (Fluent) инициализировано.")

    def _setup_ui(self):
        """Настраивает пользовательский интерфейс с Fluent Widgets."""
        logger.debug("Настройка UI главного окна (Fluent)...")

        # --- Пример: Добавление страницы "Главная" ---
        home_interface = self._create_home_interface()
        self.addSubInterface(home_interface, 'home_interface', 'Главная', icon=None) # QIcon.fromTheme('home')

        # --- Пример: Добавление страницы "Проект" ---
        project_interface = self._create_project_interface()
        self.addSubInterface(project_interface, 'project_interface', 'Проект', icon=None) # QIcon.fromTheme('folder')

        # --- Пример: Добавление страницы "Редактор" ---
        editor_interface = self._create_editor_interface()
        self.addSubInterface(editor_interface, 'editor_interface', 'Редактор', icon=None) # QIcon.fromTheme('edit')

        # --- Пример: Добавление страницы "Настройки" ---
        settings_interface = self._create_settings_interface()
        self.addSubInterface(settings_interface, 'settings_interface', 'Настройки', icon=None) # QIcon.fromTheme('settings')

        logger.debug("UI главного окна (Fluent) настроено.")

    def _create_home_interface(self) -> QWidget:
        """Создаёт интерфейс для главной страницы."""
        interface = QWidget()
        layout = QVBoxLayout(interface)
        label = SubtitleLabel('Главная страница')
        setFont(label, 24)
        layout.addWidget(label)
        layout.addWidget(QLabel('Добро пожаловать в Excel Micro DB!'))
        # Здесь можно добавить кнопки быстрого доступа, последние проекты и т.д.
        return interface

    def _create_project_interface(self) -> QWidget:
        """Создаёт интерфейс для страницы проекта."""
        interface = QWidget()
        layout = QVBoxLayout(interface)
        label = SubtitleLabel('Проект')
        setFont(label, 24)
        layout.addWidget(label)
        layout.addWidget(QLabel('Здесь будет отображаться структура проекта и его данные.'))
        # Здесь будет ProjectExplorer или аналогичный виджет
        return interface

    def _create_editor_interface(self) -> QWidget:
        """Создаёт интерфейс для страницы редактора."""
        interface = QWidget()
        layout = QVBoxLayout(interface)
        label = SubtitleLabel('Редактор')
        setFont(label, 24)
        layout.addWidget(label)
        layout.addWidget(QLabel('Здесь будет редактор листов Excel.'))
        # Здесь будет SheetEditor или аналогичный виджет
        return interface

    def _create_settings_interface(self) -> QWidget:
        """Создаёт интерфейс для страницы настроек."""
        interface = QWidget()
        layout = QVBoxLayout(interface)
        label = SubtitleLabel('Настройки')
        setFont(label, 24)
        layout.addWidget(label)
        layout.addWidget(QLabel('Здесь будут настройки приложения.'))
        # Здесь будут настройки приложения
        return interface

    # def _create_actions(self): ...
    # def _create_menus(self): ...
    # def _create_toolbars(self): ...
    # def _create_status_bar(self): ...
    # def _connect_signals(self): ...

    # --- Слоты для действий ---
    # @Slot()
    # def _on_new_project(self): ...
    # @Slot()
    # def _on_open_project(self): ...
    # @Slot()
    # def _on_analyze_excel(self): ...
    # @Slot()
    # def _on_export_to_excel(self): ...
    # @Slot(str)
    # def _on_sheet_selected(self, sheet_name: str): ...
