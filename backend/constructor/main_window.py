# backend/constructor/main_window.py
"""
Главное окно приложения Excel Micro DB.
"""

import sys
import os
from pathlib import Path
from typing import Optional, Dict, Any, List

# Импорты из PySide6
from PySide6.QtCore import Qt, Slot, Signal
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QMenu, QStatusBar, QToolBar,
    QStackedWidget, QWidget, QMessageBox, QFileDialog, QDockWidget, QLabel, QVBoxLayout
)

# Импорт внутренних модулей
# ИСПРАВЛЕНО: Импорты теперь относительные или абсолютные внутри backend
from backend.utils.logger import get_logger # <-- ИСПРАВЛЕНО: было from utils.logger
from backend.core.app_controller import create_app_controller, AppController
from backend.constructor.widgets.project_explorer import ProjectExplorer
# ИСПРАВЛЕНО: Импорт SheetEditor теперь из подпапки sheet_editor
from backend.constructor.widgets.sheet_editor.sheet_editor_widget import SheetEditor

# Встроенный WelcomeWidget для устранения ошибки импорта
class WelcomeWidget(QWidget):
    """Простая заглушка для приветственного экрана."""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        label = QLabel("Добро пожаловать в Excel Micro DB!\nПожалуйста, создайте или откройте проект.")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Улучшаем стиль метки
        label.setStyleSheet("font-size: 16px; padding: 20px;")
        layout.addWidget(label)
        self.setLayout(layout)

logger = get_logger(__name__)

class MainWindow(QMainWindow):
    """
    Главное окно приложения.
    Является центральным элементом графического интерфейса.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        logger.debug("Создание экземпляра MainWindow")
        self.setWindowTitle("Excel Micro DB")
        self.resize(1200, 800)

        # Инициализация атрибутов
        self.app_controller: Optional[AppController] = None
        self.project_explorer: Optional[ProjectExplorer] = None
        self.sheet_editor: Optional[SheetEditor] = None
        self.welcome_widget: Optional[QWidget] = None
        self.stacked_widget: Optional[QStackedWidget] = None
        self.project_explorer_dock: Optional[QDockWidget] = None

        # Создание и настройка UI
        self._setup_ui()
        self._create_actions()
        self._create_menus()
        self._create_toolbars()
        self._create_status_bar()

        # Инициализация контроллера приложения
        self._init_app_controller()

        # Показать приветственный экран
        self._show_welcome_screen()

        logger.info("MainWindow инициализировано")

    def _setup_ui(self):
        """Настройка пользовательского интерфейса."""
        # Создаем центральный стек виджетов
        self.stacked_widget = QStackedWidget()
        if self.stacked_widget:
            self.setCentralWidget(self.stacked_widget)

        # Создаем виджеты
        self.welcome_widget = WelcomeWidget()
        self.sheet_editor = SheetEditor()

        # Добавляем виджеты в стек
        if self.stacked_widget and self.welcome_widget:
            self.stacked_widget.addWidget(self.welcome_widget)
        if self.stacked_widget and self.sheet_editor:
            self.stacked_widget.addWidget(self.sheet_editor)

        # Создание и настройка Project Explorer в док-виджете
        self.project_explorer = ProjectExplorer(self) # Передаем parent

        # Создаем ОТДЕЛЬНЫЙ QDockWidget, который будет контейнером
        self.project_explorer_dock = QDockWidget("Проект", self)

        if self.project_explorer_dock:
            # Разрешаем закреплять слева и справа
            self.project_explorer_dock.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)

        # Помещаем ЭКЗЕМПЛЯР ProjectExplorer (QWidget) ВНУТРЬ QDockWidget
        if self.project_explorer_dock and self.project_explorer:
            self.project_explorer_dock.setWidget(self.project_explorer)

        # Добавляем ГОТОВЫЙ QDockWidget в главное окно
        if self.project_explorer_dock:
            self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.project_explorer_dock)

        # Подключаем сигналы
        if self.project_explorer and hasattr(self.project_explorer, 'sheet_selected'):
            sheet_selected_attr = getattr(self.project_explorer, 'sheet_selected')
            if callable(getattr(sheet_selected_attr, 'connect', None)):
                sheet_selected_attr.connect(self._on_sheet_selected)
                logger.debug("Сигнал sheet_selected успешно подключен.")
            else:
                error_msg = "Атрибут 'sheet_selected' в ProjectExplorer не является сигналом."
                logger.error(error_msg)
                QMessageBox.warning(self, "Ошибка GUI", f"{error_msg} Функциональность выбора листа может быть нарушена.")
        else:
            error_msg = "Сигнал 'sheet_selected' не найден в ProjectExplorer."
            logger.error(error_msg)

        if self.sheet_editor and self.app_controller:
            # Устанавливаем контроллер для sheet_editor
            self.sheet_editor.set_app_controller(self.app_controller)

    def _create_actions(self):
        """Создание действий (QAction) для меню и панелей инструментов."""
        self.action_new_project = QAction("&Новый проект", self)
        self.action_new_project.setShortcut("Ctrl+N")
        self.action_new_project.triggered.connect(self._on_new_project)

        self.action_open_project = QAction("&Открыть проект", self)
        self.action_open_project.setShortcut("Ctrl+O")
        self.action_open_project.triggered.connect(self._on_open_project)

        self.action_save_project = QAction("&Сохранить проект", self)
        self.action_save_project.setShortcut("Ctrl+S")
        # Сохранение проекта может быть реализовано позже или через контроллер
        # self.action_save_project.triggered.connect(self._on_save_project)

        self.action_close_project = QAction("&Закрыть проект", self)
        self.action_close_project.triggered.connect(self._on_close_project)

        self.action_exit = QAction("&Выход", self)
        self.action_exit.setShortcut("Ctrl+Q")
        self.action_exit.triggered.connect(self.close)

        self.action_analyze_excel = QAction("&Анализировать Excel-файл", self)
        self.action_analyze_excel.triggered.connect(self._on_analyze_excel)

        self.action_export_to_excel = QAction("&Экспортировать в Excel", self)
        self.action_export_to_excel.triggered.connect(self._on_export_to_excel)

    def _create_menus(self):
        """Создание меню."""
        menubar = self.menuBar()
        # В macOS меню обычно находится в строке меню системы, а не в окне.
        # Qt обрабатывает это автоматически, но проверка не повредит.
        if not menubar:
             logger.error("Не удалось получить QMenuBar")
             return


        # Меню "Файл"
        file_menu = menubar.addMenu("&Файл")
        if file_menu:
            file_menu.addAction(self.action_new_project)
            file_menu.addAction(self.action_open_project)
            file_menu.addAction(self.action_save_project)
            file_menu.addSeparator()
            file_menu.addAction(self.action_close_project)
            file_menu.addAction(self.action_exit)

        # Меню "Инструменты"
        tools_menu = menubar.addMenu("&Инструменты")
        if tools_menu:
            tools_menu.addAction(self.action_analyze_excel)
            tools_menu.addAction(self.action_export_to_excel)

    def _create_toolbars(self):
        """Создание панелей инструментов."""
        toolbar = self.addToolBar("Основные")
        if toolbar:
            toolbar.addAction(self.action_new_project)
            toolbar.addAction(self.action_open_project)
            toolbar.addAction(self.action_analyze_excel)
            toolbar.addAction(self.action_export_to_excel)

    def _create_status_bar(self):
        """Создание строки состояния."""
        self.statusBar().showMessage("Готов")

    def _init_app_controller(self):
        """Инициализация контроллера приложения."""
        logger.debug("Создание AppController")
        try:
            # Создаем контроллер
            self.app_controller = create_app_controller()
            # Инициализируем его
            if self.app_controller:
                 success = self.app_controller.initialize()
                 if success:
                     logger.info("AppController инициализирован и готов к работе")
                     # Устанавливаем контроллер для sheet_editor, если он уже создан
                     if self.sheet_editor:
                         self.sheet_editor.set_app_controller(self.app_controller)
                 else:
                     logger.error("Не удалось инициализировать AppController")
                     QMessageBox.critical(self, "Ошибка", "Не удалось инициализировать контроллер приложения.")
            else:
                logger.error("Не удалось создать AppController")
                QMessageBox.critical(self, "Ошибка", "Не удалось создать контроллер приложения.")
        except Exception as e:
            logger.error(f"Ошибка при инициализации AppController: {e}", exc_info=True)
            QMessageBox.critical(self, "Критическая ошибка", f"Ошибка при инициализации контроллера приложения:\n{e}")

    @Slot(str)
    def _on_sheet_selected(self, sheet_name: str):
        """Слот для обработки сигнала выбора листа в ProjectExplorer."""
        logger.debug(f"MainWindow получил сигнал о выборе листа: {sheet_name}")
        if not self.app_controller:
            logger.warning("AppController не инициализирован.")
            return
        if not getattr(self.app_controller, 'is_project_loaded', False):
            logger.warning("Попытка загрузить лист без загруженного проекта.")
            return

        if self.sheet_editor and self.stacked_widget:
            # Переключаемся на виджет редактора листа
            if isinstance(self.sheet_editor, QWidget):
                self.stacked_widget.setCurrentWidget(self.sheet_editor)
            else:
                logger.error("sheet_editor не является экземпляром QWidget.")
                return

            # Получаем путь к БД проекта
            db_path: Optional[str] = None
            if self.app_controller:
                project_path_attr = getattr(self.app_controller, 'project_path', None)
                if project_path_attr:
                    try:
                        if isinstance(project_path_attr, Path):
                            db_path_obj = project_path_attr / "project_data.db"
                        else: # Предполагаем строку
                            db_path_obj = Path(project_path_attr) / "project_data.db"
                        db_path = str(db_path_obj)
                        logger.debug(f"Путь к БД определен через app_controller.project_path: {db_path}")
                    except Exception as e:
                        logger.error(f"Ошибка при формировании пути к БД: {e}")

            if db_path and self.sheet_editor:
                logger.debug(f"Загрузка листа '{sheet_name}' с использованием БД: {db_path}")
                # Явно утверждаем, что sheet_editor - это SheetEditor
                assert isinstance(self.sheet_editor, SheetEditor), "sheet_editor должен быть экземпляром SheetEditor"
                self.sheet_editor.load_sheet(db_path, sheet_name)
            elif not db_path:
                error_msg = "Не удалось определить путь к БД проекта для загрузки листа."
                logger.error(error_msg)
                QMessageBox.critical(self, "Ошибка", f"{error_msg}\nУбедитесь, что проект корректно загружен.")
            else:
                logger.error("sheet_editor или stacked_widget не инициализированы при попытке загрузки листа.")
        else:
            error_msg = "sheet_editor или stacked_widget не инициализированы."
            logger.error(error_msg)

    def _show_welcome_screen(self):
        """Показывает приветственный экран."""
        if self.stacked_widget and self.welcome_widget:
            if isinstance(self.welcome_widget, QWidget):
                self.stacked_widget.setCurrentWidget(self.welcome_widget)
            else:
                 logger.warning("welcome_widget не является экземпляром QWidget.")
        else:
             logger.debug("Невозможно показать приветственный экран: stacked_widget или welcome_widget отсутствуют.")

    def _on_new_project(self):
        """Обработчик действия 'Новый проект'."""
        logger.info("Начало создания нового проекта")
        project_path = QFileDialog.getExistingDirectory(self, "Выберите директорию для нового проекта")
        if project_path:
            if self.app_controller:
                success = self.app_controller.create_project(project_path)
                if success:
                    logger.info(f"Проект успешно создан в: {project_path}")
                    self._load_project(project_path)
                else:
                    logger.error(f"Не удалось создать проект в: {project_path}")
                    QMessageBox.critical(self, "Ошибка", "Не удалось создать проект.")
            else:
                logger.error("AppController не инициализирован.")
                QMessageBox.critical(self, "Ошибка", "Контроллер приложения не доступен.")

    def _on_open_project(self):
        """Обработчик действия 'Открыть проект'."""
        logger.info("Начало открытия проекта")
        project_path = QFileDialog.getExistingDirectory(self, "Выберите директорию проекта")
        if project_path:
            self._load_project(project_path)

    def _load_project(self, project_path: str):
        """Загружает проект по указанному пути."""
        if not self.app_controller:
            logger.error("AppController не инициализирован.")
            QMessageBox.critical(self, "Ошибка", "Контроллер приложения не доступен.")
            return

        success = self.app_controller.load_project(project_path)
        if success:
            logger.info(f"Проект успешно загружен из: {project_path}")
            # Обновляем ProjectExplorer
            if self.project_explorer and self.app_controller:
                # Получаем project_data из app_controller
                project_data: Optional[Dict[str, Any]] = getattr(self.app_controller, 'current_project', None)
                if project_data is None:
                     logger.error("AppController.current_project отсутствует или None.")
                     QMessageBox.warning(self, "Предупреждение", "Не удалось получить данные проекта из AppController.")
                     return

                # Получаем путь к БД проекта
                db_path: Optional[str] = None
                if self.app_controller:
                    project_path_attr = getattr(self.app_controller, 'project_path', None)
                    if project_path_attr:
                        try:
                            if isinstance(project_path_attr, Path):
                                db_path_obj = project_path_attr / "project_data.db"
                            else: # Предполагаем строку
                                db_path_obj = Path(project_path_attr) / "project_data.db"
                            db_path = str(db_path_obj)
                            logger.debug(f"Путь к БД для ProjectExplorer: {db_path}")
                        except Exception as e:
                            logger.error(f"Ошибка при формировании пути к БД для ProjectExplorer: {e}")
                            QMessageBox.warning(self, "Предупреждение", f"Ошибка при формировании пути к БД: {e}")
                            return # Прерываем, если не можем получить путь к БД

                # Проверяем, получены ли оба необходимых аргумента
                if project_data is not None and db_path:
                    try:
                        self.project_explorer.load_project(project_data, db_path)
                        logger.debug("ProjectExplorer успешно загрузил структуру проекта.")
                    except Exception as e:
                         error_msg = f"Неожиданная ошибка при вызове load_project в ProjectExplorer: {e}"
                         logger.error(error_msg, exc_info=True)
                         QMessageBox.warning(self, "Предупреждение", error_msg)
                else:
                    missing = []
                    if project_data is None:
                        missing.append("project_data")
                    if not db_path:
                        missing.append("db_path")
                    error_msg = f"Не удалось получить необходимые данные для ProjectExplorer: {', '.join(missing)}."
                    logger.error(error_msg)
                    QMessageBox.warning(self, "Предупреждение", error_msg)

            self._show_welcome_screen()
        else:
            logger.error(f"Не удалось загрузить проект из: {project_path}")
            QMessageBox.critical(self, "Ошибка", "Не удалось загрузить проект.")

    def _on_close_project(self):
        """Обработчик действия 'Закрыть проект'."""
        if not self.app_controller:
             logger.info("AppController не инициализирован при попытке закрыть проект.")
             return
        if not getattr(self.app_controller, 'is_project_loaded', False):
            logger.info("Попытка закрыть проект, но проект не был загружен.")
            return

        reply = QMessageBox.question(self, 'Подтверждение', 'Вы уверены, что хотите закрыть проект?',
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            logger.info("Пользователь подтвердил закрытие проекта.")
            if self.project_explorer:
                self.project_explorer.clear_project()
            if self.sheet_editor:
                self.sheet_editor.clear_sheet()
            self._show_welcome_screen()
            logger.info("Проект закрыт.")

    def _on_analyze_excel(self):
        """Обработчик действия 'Анализировать Excel-файл'."""
        if not self.app_controller:
             QMessageBox.critical(self, "Ошибка", "Контроллер приложения не доступен.")
             return
        if not getattr(self.app_controller, 'is_project_loaded', False):
            QMessageBox.warning(self, "Предупреждение", "Сначала необходимо открыть или создать проект.")
            return

        logger.info("Начало анализа Excel-файла")
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel-файл для анализа", "",
                                                   "Excel Files (*.xlsx *.xlsm)")
        if file_path:
            success = self.app_controller.analyze_excel_file(file_path)
            if success:
                logger.info(f"Файл {file_path} успешно проанализирован.")
                QMessageBox.information(self, "Успех", "Файл успешно проанализирован.")
                # Обновляем ProjectExplorer ПОСЛЕ АНАЛИЗА
                if self.project_explorer and self.app_controller:
                     # Получаем project_data из app_controller
                     project_data: Optional[Dict[str, Any]] = getattr(self.app_controller, 'current_project', None)
                     if project_data is None:
                          logger.error("AppController.current_project отсутствует или None (после анализа).")
                          QMessageBox.warning(self, "Предупреждение", "Не удалось получить данные проекта из AppController (после анализа).")
                          return

                     # Получаем путь к БД проекта (он должен быть уже установлен после анализа/загрузки)
                     db_path: Optional[str] = None
                     if self.app_controller:
                        project_path_attr = getattr(self.app_controller, 'project_path', None)
                        if project_path_attr:
                            try:
                                if isinstance(project_path_attr, Path):
                                    db_path_obj = project_path_attr / "project_data.db"
                                else: # Предполагаем строку
                                    db_path_obj = Path(project_path_attr) / "project_data.db"
                                db_path = str(db_path_obj)
                                logger.debug(f"Путь к БД для ProjectExplorer (после анализа): {db_path}")
                            except Exception as e:
                                logger.error(f"Ошибка при формировании пути к БД для ProjectExplorer (после анализа): {e}")
                                QMessageBox.warning(self, "Предупреждение", f"Ошибка при формировании пути к БД (после анализа): {e}")
                                return # Прерываем, если не можем получить путь к БД

                     # Проверяем, получены ли оба необходимых аргумента
                     if project_data is not None and db_path:
                        try:
                            self.project_explorer.load_project(project_data, db_path)
                            logger.debug("ProjectExplorer успешно обновлен после анализа.")
                        except Exception as e:
                             error_msg = f"Неожиданная ошибка при обновлении ProjectExplorer после анализа: {e}"
                             logger.error(error_msg, exc_info=True)
                             QMessageBox.warning(self, "Предупреждение", error_msg)
                     else:
                         missing = []
                         if project_data is None:
                             missing.append("project_data")
                         if not db_path:
                             missing.append("db_path")
                         error_msg = f"Не удалось получить необходимые данные для ProjectExplorer после анализа: {', '.join(missing)}."
                         logger.error(error_msg)
                         QMessageBox.warning(self, "Предупреждение", error_msg)
            else:
                logger.error(f"Не удалось проанализировать файл {file_path}.")
                QMessageBox.critical(self, "Ошибка", "Не удалось проанализировать файл.")

    def _on_export_to_excel(self):
        """Обработчик действия 'Экспортировать в Excel'."""
        if not self.app_controller:
             QMessageBox.critical(self, "Ошибка", "Контроллер приложения не доступен.")
             return
        if not getattr(self.app_controller, 'is_project_loaded', False):
            QMessageBox.warning(self, "Предупреждение", "Сначала необходимо открыть или создать проект.")
            return

        logger.info("Начало экспорта проекта в Excel")
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", "",
                                                   "Excel Files (*.xlsx)")
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            # Сигнатура: export_results(self, export_type: str, output_path: str) -> bool:
            success = self.app_controller.export_results('excel', file_path)
            if success:
                logger.info(f"Проект успешно экспортирован в: {file_path}")
                QMessageBox.information(self, "Успех", f"Проект успешно экспортирован в:\n{file_path}")
            else:
                logger.error(f"Не удалось экспортировать проект в: {file_path}")
                QMessageBox.critical(self, "Ошибка", "Не удалось экспортировать проект.")

    def closeEvent(self, event):
        """Обработчик события закрытия окна."""
        logger.info("Получен запрос на закрытие главного окна")
        if self.app_controller:
            self.app_controller.shutdown()
        event.accept()
        logger.info("Главное окно закрыто")
