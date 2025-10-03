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
from backend.analyzer.logic_documentation import analyze_excel_file
from backend.storage.base import ProjectDBStorage
from backend.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter
from backend.constructor.widgets.simple_gui.welcome_widget import WelcomeWidget
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
        self.project_db_path: Optional[str] = None
        self.db_storage: Optional[ProjectDBStorage] = None
        self.sheet_names: List[str] = []
        
        # Виджеты
        self.welcome_widget: Optional[WelcomeWidget] = None
        self.sheet_editor: Optional[SimpleSheetEditor] = None
        self.stacked_widget: Optional[QStackedWidget] = None
        self.file_explorer: Optional[ExcelFileExplorer] = None
        
        # Настройка UI
        self._setup_ui()
        self._create_actions()
        self._create_menus()
        self._create_toolbars()
        self._create_status_bar()
        
        # Показываем приветственный экран
        self._show_welcome_screen()
        
        logger.info("SimpleMainWindow инициализировано")
    
    def _setup_ui(self):
        """Настройка пользовательского интерфейса."""
        # Центральный стек виджетов
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        
        # Создаем виджеты
        self.welcome_widget = WelcomeWidget()
        self.sheet_editor = SimpleSheetEditor()
        
        # Добавляем в стек
        self.stacked_widget.addWidget(self.welcome_widget)
        self.stacked_widget.addWidget(self.sheet_editor)
        
        # Создаем проводник файла
        self.file_explorer = ExcelFileExplorer(self)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.file_explorer)
        
        # Подключаем сигналы
        self.welcome_widget.import_requested.connect(self._on_import_excel)
        self.file_explorer.sheet_selected.connect(self._on_sheet_selected)
    
    def _create_actions(self):
        """Создание действий для меню и панелей инструментов."""
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
        file_menu.addAction(self.action_import_excel)
        file_menu.addAction(self.action_export_excel)
        file_menu.addSeparator()
        file_menu.addAction(self.action_close_file)
        file_menu.addAction(self.action_exit)
    
    def _create_toolbars(self):
        """Создание панелей инструментов."""
        toolbar = self.addToolBar("Основные")
        toolbar.addAction(self.action_import_excel)
        toolbar.addAction(self.action_export_excel)
    
    def _create_status_bar(self):
        """Создание строки состояния."""
        self.statusBar().showMessage("Готов")
    
    def _show_welcome_screen(self):
        """Показывает приветственный экран."""
        if self.stacked_widget and self.welcome_widget:
            self.stacked_widget.setCurrentWidget(self.welcome_widget)
            # Убраны строки, которые принудительно выключали кнопки
            # self.action_export_excel.setEnabled(False)
            # self.action_close_file.setEnabled(False)
            if self.file_explorer:
                self.file_explorer.clear_file()
    
    def _on_import_excel(self):
        """Обработчик импорта Excel-файла."""
        logger.info("Начало импорта Excel-файла")
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Выберите Excel-файл для анализа", 
            "", 
            "Excel Files (*.xlsx *.xlsm)"
        )
        
        if not file_path:
            return
        
        # Спрашиваем у пользователя папку проекта
        project_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите директорию для проекта",
            "",
        )
        if not project_path:
            logger.info("Пользователь отменил выбор папки проекта.")
            return
        
        project_path = Path(project_path).resolve()
        
        try:
            # Анализируем файл
            logger.info(f"Анализ файла: {file_path}")
            analysis_results = analyze_excel_file(file_path)
            
            if not analysis_results or "sheets" not in analysis_results:
                raise Exception("Анализ не вернул ожидаемые результаты")
            
            # Получаем имена листов
            self.sheet_names = [sheet["name"] for sheet in analysis_results["sheets"]]
            if not self.sheet_names:
                raise Exception("В файле не найдено листов")
            
            # Устанавливаем пути к проекту и БД
            self.project_path = str(project_path)
            self.project_db_path = str(project_path / "project_data.db")
            self.db_storage = ProjectDBStorage(self.project_db_path)
            
            # Инициализируем схему БД
            if not self.db_storage.initialize_project_tables():
                raise Exception("Не удалось инициализировать БД")
            
            # Сохраняем результаты анализа в БД (с подключением внутри)
            self._save_analysis_to_db(analysis_results)
            
            # Обновляем состояние
            self.current_excel_file = file_path
            
            # Обновляем UI
            if self.file_explorer:
                self.file_explorer.load_excel_file(file_path, self.sheet_names)
            self.action_export_excel.setEnabled(True)
            self.action_close_file.setEnabled(True)
            self._show_welcome_screen()  # Остаемся на приветственном экране
            
            logger.info(f"Файл успешно проанализирован: {file_path}, БД: {self.project_db_path}")
            QMessageBox.information(self, "Успех", "Файл успешно проанализирован!")
            
        except Exception as e:
            logger.error(f"Ошибка при импорте/анализе файла: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось проанализировать файл:\n{e}")
            self._cleanup_resources()
    
    def _save_analysis_to_db(self, analysis_results: Dict[str, Any]):
        """Сохраняет результаты анализа в БД проекта."""
        if not self.db_storage:
            raise Exception("БД не инициализирована")
        
        # Подключаемся к БД перед сохранением
        if not self.db_storage.connect():
            raise Exception("Не удалось подключиться к БД для сохранения анализа")
        
        try:
            # Сохраняем данные для каждого листа
            for sheet_data in analysis_results.get("sheets", []):
                sheet_name = sheet_data["name"]
                logger.info(f"Сохранение данных для листа: {sheet_name}")
                
                # Получаем или создаем ID листа (теперь connection гарантированно есть)
                sheet_id = self._get_or_create_sheet_id(sheet_name)
                if sheet_id is None:
                    logger.error(f"Не удалось получить/создать ID для листа '{sheet_name}'. Пропущен.")
                    continue
                
                # Сохраняем метаданные листа
                metadata_to_save = {
                    "max_row": sheet_data.get("max_row"),
                    "max_column": sheet_data.get("max_column"),
                    "merged_cells": sheet_data.get("merged_cells", [])
                }
                if not self.db_storage.save_sheet_metadata(sheet_name, metadata_to_save):
                    logger.warning(f"Не удалось сохранить метаданные для листа '{sheet_name}'.")
                
                # Сохраняем объединенные ячейки
                merged_cells_list = sheet_data.get("merged_cells", [])
                if merged_cells_list:
                    if not self.db_storage.save_sheet_merged_cells(sheet_id, merged_cells_list):
                        logger.error(f"Не удалось сохранить объединенные ячейки для листа '{sheet_name}' (ID: {sheet_id}).")
                
                # Сохраняем "сырые данные"
                if not self.db_storage.save_sheet_raw_data(sheet_name, sheet_data.get("raw_data", [])):
                    logger.error(f"Не удалось сохранить 'сырые данные' для листа '{sheet_name}'.")
                
                # Сохраняем формулы
                if not self.db_storage.save_sheet_formulas(sheet_id, sheet_data.get("formulas", [])):
                    logger.error(f"Не удалось сохранить формулы для листа '{sheet_name}' (ID: {sheet_id}).")
                
                # Сохраняем стили
                if not self.db_storage.save_sheet_styles(sheet_id, sheet_data.get("styles", [])):
                    logger.error(f"Не удалось сохранить стили для листа '{sheet_name}' (ID: {sheet_id}).")
                
                # Сохраняем диаграммы
                if not self.db_storage.save_sheet_charts(sheet_id, sheet_data.get("charts", [])):
                    logger.error(f"Не удалось сохранить диаграммы для листа '{sheet_name}' (ID: {sheet_id}).")
        finally:
            # Отключаемся от БД после сохранения
            self.db_storage.disconnect()
    
    def _get_or_create_sheet_id(self, sheet_name: str) -> Optional[int]:
        """Получает ID листа из БД или создает новую запись."""
        if not self.db_storage or not self.db_storage.connection:
            logger.error("Нет подключения к БД для получения/создания sheet_id.")
            return None
        
        try:
            cursor = self.db_storage.connection.cursor()
            project_id = 1  # Для MVP используем ID 1
            
            # Используем правильное имя столбца 'sheet_id'
            cursor.execute("SELECT sheet_id FROM sheets WHERE project_id = ? AND name = ?", (project_id, sheet_name))
            result = cursor.fetchone()
            if result:
                return result[0]
            else:
                cursor.execute("INSERT INTO sheets (project_id, name) VALUES (?, ?)", (project_id, sheet_name))
                self.db_storage.connection.commit()
                return cursor.lastrowid
        except Exception as e:
            logger.error(f"Ошибка при получении/создании sheet_id для '{sheet_name}': {e}", exc_info=True)
            return None
    
    @Slot(str)
    def _on_sheet_selected(self, sheet_name: str):
        """Обработчик выбора листа."""
        logger.debug(f"Выбран лист: {sheet_name}")
        if not self.project_db_path:
            logger.warning("Нет активной БД")
            return
        
        if self.sheet_editor and self.stacked_widget:
            self.stacked_widget.setCurrentWidget(self.sheet_editor)
            self.sheet_editor.load_sheet(self.project_db_path, sheet_name)
    
    def _on_export_excel(self):
        """Обработчик экспорта в Excel."""
        if not self.project_db_path or not self.current_excel_file:
            QMessageBox.warning(self, "Предупреждение", "Нет активного файла для экспорта")
            return
        
        # Предлагаем сохранить с именем оригинального файла + "_recreated"
        original_path = Path(self.current_excel_file)
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
            success = export_project_xlsxwriter(self.project_db_path, output_path)
            
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
            'Вы уверены, что хотите закрыть текущий файл?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self._cleanup_resources()
            self._show_welcome_screen()
            logger.info("Файл закрыт")
    
    def _cleanup_resources(self):
        """Очищает ресурсы."""
        # Закрываем соединение с БД
        if self.db_storage:
            self.db_storage.disconnect()
            self.db_storage = None
        
        # Сбрасываем состояние
        self.current_excel_file = None
        self.project_path = None
        self.project_db_path = None
        self.sheet_names = []
    
    def closeEvent(self, event):
        """Обработчик закрытия окна."""
        logger.info("Закрытие приложения")
        self._cleanup_resources()
        event.accept()