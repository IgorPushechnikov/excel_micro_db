# backend/constructor/widgets/simple_gui/main_window.py
"""
Главное окно упрощённого GUI Excel Micro DB.
Работает с одним Excel-файлом напрямую без концепции "проекта".
"""
import sys
import os
from pathlib import Path
from typing import Optional, Dict, Any, List

# Импорт Qt
from PySide6.QtWidgets import (
    QMainWindow, QStackedWidget, QStatusBar, QToolBar,
    QWidget, QMessageBox, QFileDialog
)
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QAction

# Импорт внутренних модулей
from backend.utils.logger import get_logger
# УДАЛЯЕМ: from backend.analyzer.logic_documentation import analyze_excel_file
# УДАЛЯЕМ: from backend.storage.base import ProjectDBStorage
# УДАЛЯЕМ: from backend.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter
from backend.core.app_controller import create_app_controller, AppController
# from backend.constructor.widgets.simple_gui.welcome_widget import WelcomeWidget # Больше не используется
from backend.constructor.widgets.simple_gui.file_explorer import ExcelFileExplorer
from backend.constructor.widgets.simple_gui.simple_sheet_editor import SimpleSheetEditor

logger = get_logger(__name__)


class SimpleMainWindow(QMainWindow):
    """Упрощенное главное окно приложения."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        logger.debug("Создание экземпляра SimpleMainWindow")
        self.setWindowTitle("Excel Micro DB - Simple")
        self.resize(1200, 800)
        
        # Атрибуты для работы с файлом
        self.current_excel_file: Optional[str] = None
        self.project_path: Optional[str] = None
        # УДАЛЯЕМ: self.project_db_path: Optional[str] = None
        # УДАЛЯЕМ: self.db_storage: Optional[ProjectDBStorage] = None
        self.app_controller: Optional[AppController] = None  # <-- НОВОЕ
        self.sheet_names: List[str] = []
        
        # Виджеты
        # self.welcome_widget: Optional[WelcomeWidget] = None # Больше не используется
        self.sheet_editor: Optional[SimpleSheetEditor] = None
        self.stacked_widget: Optional[QStackedWidget] = None
        self.file_explorer: Optional[ExcelFileExplorer] = None
        
        # Настройка UI
        self._setup_ui()
        self._create_actions()
        self._create_menus()
        self._create_toolbars()
        self._create_status_bar()
        
        # --- ИЗМЕНЕНИЕ: Убран вызов _show_welcome_screen ---
        # Показываем редактор с пустой таблицей
        if self.stacked_widget and self.sheet_editor:
            self.stacked_widget.setCurrentWidget(self.sheet_editor)
        # -----------------------------------------------
        
        logger.info("SimpleMainWindow инициализировано")
    
    def _setup_ui(self):
        """Настройка пользовательского интерфейса."""
        # Центральный стек виджетов
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        
        # Создаем виджеты
        # self.welcome_widget = WelcomeWidget() # Больше не используется
        self.sheet_editor = SimpleSheetEditor()
        
        # Добавляем в стек
        # self.stacked_widget.addWidget(self.welcome_widget) # Больше не используется
        self.stacked_widget.addWidget(self.sheet_editor)
        
        # Создаем проводник файла
        self.file_explorer = ExcelFileExplorer(self)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.file_explorer)
        
        # --- ИЗМЕНЕНИЕ: Убрана связь с welcome_widget ---
        # self.welcome_widget.import_requested.connect(self._on_import_excel)
        # -------------------------------------------
        
        # Подключаем сигналы, если нужно
        # self.file_explorer.sheet_selected.connect(self._on_sheet_selected)
        # Пока что не подключаем, так как нет файла

    def _create_actions(self):
        """Создание действий для меню и панелей инструментов."""
        # --- ИЗМЕНЕНИЕ: Добавлены действия для нового/открытия проекта ---
        self.action_new_project = QAction("&Новый проект", self)
        self.action_new_project.setShortcut("Ctrl+N")
        self.action_new_project.triggered.connect(self._on_new_project)
        
        self.action_open_project = QAction("&Открыть проект", self)
        self.action_open_project.setShortcut("Ctrl+Shift+O")
        self.action_open_project.triggered.connect(self._on_open_project)
        # -------------------------------------------------------------

        self.action_import_excel = QAction("&Импорт Excel-файла", self)
        self.action_import_excel.setShortcut("Ctrl+O")
        self.action_import_excel.triggered.connect(self._on_import_excel)
        
        self.action_export_excel = QAction("&Экспорт в Excel", self)
        self.action_export_excel.setShortcut("Ctrl+S")
        self.action_export_excel.triggered.connect(self._on_export_excel)
        self.action_export_excel.setEnabled(False)  # Пока нет файла
        
        self.action_close_file = QAction("&Закрыть файл", self)
        self.action_close_file.triggered.connect(self._on_close_file)
        self.action_close_file.setEnabled(False)

        self.action_exit = QAction("&Выход", self)
        self.action_exit.setShortcut("Ctrl+Q")
        self.action_exit.triggered.connect(self.close)
    
    def _create_menus(self):
        """Создание меню."""
        menubar = self.menuBar()

        # Меню "Файл"
        file_menu = menubar.addMenu("&Файл")
        # --- ИЗМЕНЕНИЕ: Добавлены пункты меню для нового/открытия проекта ---
        file_menu.addAction(self.action_new_project)
        file_menu.addAction(self.action_open_project)
        file_menu.addSeparator()
        # -------------------------------------------------------------
        file_menu.addAction(self.action_import_excel)
        file_menu.addAction(self.action_export_excel)
        file_menu.addSeparator()
        file_menu.addAction(self.action_close_file)
        file_menu.addAction(self.action_exit)
    
    def _create_toolbars(self):
        """Создание панелей инструментов."""
        toolbar = self.addToolBar("Основные")
        # --- ИЗМЕНЕНИЕ: Добавлены кнопки для нового/открытия проекта ---
        toolbar.addAction(self.action_new_project)
        toolbar.addAction(self.action_open_project)
        toolbar.addSeparator()
        # -------------------------------------------------------------
        toolbar.addAction(self.action_import_excel)
        toolbar.addAction(self.action_export_excel)
    
    def _create_status_bar(self):
        """Создание строки состояния."""
        self.statusBar().showMessage("Готов")
    
    # --- ИЗМЕНЕНИЕ: Убран метод _show_welcome_screen ---
    # -----------------------------------------------

    # --- ИЗМЕНЕНИЕ: Убран метод _load_first_sheet ---
    # ----------------------------------------------

    # --- ИЗМЕНЕНИЕ: Новые методы для управления проектом ---
    def _on_new_project(self):
        """Обработчик создания нового проекта."""
        logger.info("Начало создания нового проекта")
        # Спрашиваем у пользователя папку проекта
        project_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите директорию для нового проекта",
            "",
        )
        if not project_path:
            logger.info("Пользователь отменил выбор папки проекта.")
            return

        try:
            # Создаем контроллер и инициализируем проект
            self.project_path = project_path
            self.app_controller = create_app_controller(self.project_path)

            # Проверяем, что app_controller инициализирован
            if not self.app_controller:
                raise Exception("AppController не инициализирован")

            # Создаём структуру проекта
            create_success = self.app_controller.create_project(self.project_path)
            if not create_success:
                raise Exception("Не удалось создать структуру проекта")

            # Явно загружаем созданный проект, чтобы установить storage
            load_success = self.app_controller.load_project(self.project_path)
            if not load_success:
                raise Exception("Не удалось загрузить только что созданный проект")

            # Обновляем состояние GUI
            self._update_gui_after_project_load()

            QMessageBox.information(self, "Успех", f"Новый проект создан в:\n{self.project_path}\n\nТеперь вы можете импортировать Excel-файл.")

        except Exception as e:
            logger.error(f"Ошибка при создании нового проекта: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось создать новый проект:\n{e}")
            self._cleanup_resources()

    def _on_open_project(self):
        """Обработчик открытия существующего проекта."""
        logger.info("Начало открытия существующего проекта")
        # Спрашиваем у пользователя папку проекта
        project_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите директорию проекта",
            "",
        )
        if not project_path:
            logger.info("Пользователь отменил выбор папки проекта.")
            return

        try:
            # Создаем контроллер и загружаем проект
            self.project_path = project_path
            self.app_controller = create_app_controller(self.project_path)

            # Проверяем, что app_controller инициализирован
            if not self.app_controller:
                raise Exception("AppController не инициализирован")

            if not self.app_controller.load_project(self.project_path):
                raise Exception("Не удалось загрузить проект")

            # Обновляем состояние GUI
            self._update_gui_after_project_load()

            # Загружаем имена листов из контроллера
            self._load_sheet_names_from_controller()

            QMessageBox.information(self, "Успех", f"Проект открыт из:\n{self.project_path}")

        except Exception as e:
            logger.error(f"Ошибка при открытии проекта: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть проект:\n{e}")
            self._cleanup_resources()


    def _update_gui_after_project_load(self):
        """Обновляет состояние GUI после загрузки проекта."""
        # --- ИЗМЕНЕНИЕ: Проверка на None для project_path ---
        if self.file_explorer and self.project_path:
            # Передаём путь к папке проекта, так как имя Excel-файла может быть неизвестно
            # или можно передать имя БД, если file_explorer это поддерживает.
            # Для совместимости с текущим file_explorer, передадим путь к БД как "имя файла".
            # Альтернатива: обновить file_explorer.load_excel_file, чтобы он принимал Optional[str]
            # и обрабатывал случай, когда file_path - None.
            self.file_explorer.load_excel_file(self.project_path, self.sheet_names)
        # -----------------------------------------------
        self.action_export_excel.setEnabled(True)
        self.action_close_file.setEnabled(True)
        # Возможно, нужно обновить статусбар или другие элементы

    def _load_sheet_names_from_controller(self):
        """Загружает имена листов из AppController."""
        if not self.app_controller:
            logger.error("Нет активного AppController для загрузки имён листов.")
            return

        try:
            self.sheet_names = self.app_controller.get_sheet_names()
            logger.info(f"Загружено {len(self.sheet_names)} имён листов через AppController.")

            # Обновляем проводник
            if self.file_explorer and self.project_path:
                self.file_explorer.load_excel_file(self.project_path, self.sheet_names)

        except Exception as e:
            logger.error(f"Ошибка при загрузке имён листов из контроллера: {e}", exc_info=True)

    # --- Возвращаем старые методы, но адаптируем под новую логику ---
    def _on_import_excel(self):
        """Обработчик импорта Excel-файла."""
        if not self.project_path or not self.app_controller:
            QMessageBox.warning(self, "Предупреждение", "Пожалуйста, сначала создайте или откройте проект.")
            return

        logger.info("Начало импорта Excel-файла")
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Выберите Excel-файл для анализа", 
            "", 
            "Excel Files (*.xlsx *.xlsm)"
        )
        
        if not file_path:
            return

        try:
            # Вызываем анализ через контроллер
            success = self.app_controller.analyze_excel_file(file_path)
            
            if not success:
                raise Exception("Анализ Excel-файла завершился с ошибкой")

            # Обновляем состояние
            self.current_excel_file = file_path
            
            # Обновляем UI
            self._load_sheet_names_from_controller()  # Перезагружаем список листов

            if self.file_explorer and self.current_excel_file:
                self.file_explorer.load_excel_file(self.current_excel_file, self.sheet_names)

            QMessageBox.information(self, "Успех", "Файл успешно проанализирован!")

        except Exception as e:
            logger.error(f"Ошибка при импорте/анализе файла: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось проанализировать файл:\n{e}")

    # _save_analysis_to_db удален, логика в AppController
    # _get_or_create_sheet_id удален, логика в AppController

    # _load_first_sheet удален, логика в _on_sheet_selected
    
    @Slot(str)
    def _on_sheet_selected(self, sheet_name: str):
        """Обработчик выбора листа."""
        logger.debug(f"Выбран лист: {sheet_name}")
        if not self.project_path or not self.app_controller:
            logger.warning("Нет активного AppController")
            return

        if self.sheet_editor and self.stacked_widget:
            self.stacked_widget.setCurrentWidget(self.sheet_editor)
            # Загружаем данные через контроллер
            try:
                # get_sheet_data возвращает Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]
                raw_data_list, styles_list = self.app_controller.get_sheet_data(sheet_name)
                # Передаем в редактор (предполагается, что SimpleSheetEditor будет обновлен)
                self.sheet_editor.load_sheet(sheet_name, raw_data_list, styles_list)
            except Exception as e:
                logger.error(f"Ошибка при загрузке данных листа '{sheet_name}': {e}", exc_info=True)
                QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить данные листа:\n{e}")
        else:
            logger.warning("sheet_editor или stacked_widget не инициализированы")
    
    def _on_export_excel(self):
        """Обработчик экспорта в Excel."""
        if not self.project_path or not self.app_controller:
            QMessageBox.warning(self, "Предупреждение", "Нет активного файла для экспорта")
            return
        
        # Предлагаем сохранить с именем оригинального файла + "_recreated"
        original_path = Path(self.current_excel_file) if self.current_excel_file else Path(self.project_path)
        default_output = original_path.parent / f"{original_path.stem}_recreated{original_path.suffix}"
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить как",
            str(default_output),
            "Excel Files (*.xlsx)"
        )
        
        if not output_path:
            return
        
        if not output_path.lower().endswith('.xlsx'):
            output_path += '.xlsx'

        try:
            logger.info(f"Экспорт в файл: {output_path}")
            # Вызываем экспорт через контроллер
            success = self.app_controller.export_results(export_type="excel", output_path=output_path)
            
            if success:
                logger.info("Экспорт завершен успешно")
                QMessageBox.information(self, "Успех", f"Файл успешно экспортирован:\n{output_path}")
            else:
                raise Exception("Экспорт завершился с ошибкой")
                
        except Exception as e:
            logger.error(f"Ошибка при экспорте: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать файл:\n{e}")
    
    def _on_close_file(self):
        """Обработчик закрытия файла."""
        reply = QMessageBox.question(
            self, 
            'Подтверждение', 
            'Вы уверены, что хотите закрыть текущий проект?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self._cleanup_resources()
            # После закрытия показываем пустой редактор
            if self.stacked_widget and self.sheet_editor:
                self.stacked_widget.setCurrentWidget(self.sheet_editor)
                self.sheet_editor.clear_sheet() # Очищаем отображение
            logger.info("Проект закрыт")
    
    def _cleanup_resources(self):
        """Очищает ресурсы."""
        # Закрываем проект через контроллер
        if self.app_controller:
            self.app_controller.close_project()
            self.app_controller = None
        
        # Сбрасываем состояние
        self.current_excel_file = None
        self.project_path = None
        # self.project_db_path = None # УДАЛЯЕМ
        self.sheet_names = []
        # Сбрасываем состояние UI
        if self.file_explorer:
            self.file_explorer.clear_file()
        self.action_export_excel.setEnabled(False)
        self.action_close_file.setEnabled(False)
    
    def closeEvent(self, event):
        """Обработчик закрытия окна."""
        logger.info("Закрытие приложения")
        self._cleanup_resources()
        event.accept()
        