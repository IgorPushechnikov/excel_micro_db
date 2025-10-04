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
        self.project_db_path: Optional[str] = None
        self.db_storage: Optional[ProjectDBStorage] = None
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

        project_path = Path(project_path).resolve()
        db_path = project_path / "project_data.db"

        # Проверяем, существует ли БД
        if db_path.exists():
            reply = QMessageBox.question(
                self,
                'Подтверждение',
                f'Файл БД {db_path} уже существует. Перезаписать его (это сотрёт все данные)?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                logger.info("Пользователь отменил перезапись существующего файла БД.")
                return
            else:
                # Удаляем существующую БД
                try:
                    db_path.unlink()
                    logger.info(f"Существующая БД {db_path} удалена.")
                except OSError as e:
                    logger.error(f"Не удалось удалить существующую БД {db_path}: {e}")
                    QMessageBox.critical(self, "Ошибка", f"Не удалось удалить старый файл БД:\n{e}")
                    return

        # Инициализируем новую БД
        try:
            self.project_path = str(project_path)
            self.project_db_path = str(db_path)
            self.db_storage = ProjectDBStorage(self.project_db_path)

            if not self.db_storage.initialize_project_tables():
                raise Exception("Не удалось инициализировать БД")

            logger.info(f"Новый проект создан: {self.project_path}, БД: {self.project_db_path}")
            QMessageBox.information(self, "Успех", f"Новый проект создан в:\n{self.project_path}\n\nТеперь вы можете импортировать Excel-файл.")

            # Обновляем состояние GUI
            self._update_gui_after_project_load()

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

        project_path = Path(project_path).resolve()
        db_path = project_path / "project_data.db"

        # Проверяем, существует ли БД
        if not db_path.exists():
            logger.error(f"Файл БД {db_path} не найден.")
            QMessageBox.critical(self, "Ошибка", f"Файл БД проекта не найден:\n{db_path}")
            return

        # Загружаем существующую БД
        try:
            self.project_path = str(project_path)
            self.project_db_path = str(db_path)
            self.db_storage = ProjectDBStorage(self.project_db_path)

            # Проверяем, инициализирована ли БД
            if not self.db_storage.connect():
                 raise Exception("Не удалось подключиться к БД")
            # Простая проверка наличия таблиц (можно улучшить)
            cursor = self.db_storage.connection.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='project_info';")
            if not cursor.fetchone():
                logger.error(f"Файл БД {db_path} не содержит корректных таблиц проекта.")
                self.db_storage.disconnect()
                raise Exception("Файл БД не содержит корректных таблиц проекта")
            self.db_storage.disconnect()

            logger.info(f"Существующий проект открыт: {self.project_path}, БД: {self.project_db_path}")
            QMessageBox.information(self, "Успех", f"Проект открыт из:\n{self.project_path}")

            # Обновляем состояние GUI
            self._update_gui_after_project_load()

            # Загружаем имена листов из БД
            self._load_sheet_names_from_db()

        except Exception as e:
            logger.error(f"Ошибка при открытии проекта: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть проект:\n{e}")
            self._cleanup_resources()

    def _update_gui_after_project_load(self):
        """Обновляет состояние GUI после загрузки проекта."""
        if self.file_explorer:
            # Пока что передаём пустой список, если не загружены листы
            self.file_explorer.load_excel_file(self.project_path, self.sheet_names)
        self.action_export_excel.setEnabled(True)
        self.action_close_file.setEnabled(True)
        # Возможно, нужно обновить статусбар или другие элементы

    def _load_sheet_names_from_db(self):
        """Загружает имена листов из БД."""
        if not self.db_storage or not self.project_db_path:
            logger.error("Нет подключения к БД для загрузки имён листов.")
            return

        try:
            if not self.db_storage.connect():
                raise Exception("Не удалось подключиться к БД для загрузки имён листов")

            cursor = self.db_storage.connection.cursor()
            cursor.execute("SELECT name FROM sheets ORDER BY name;")
            rows = cursor.fetchall()
            self.sheet_names = [row[0] for row in rows]
            logger.info(f"Загружено {len(self.sheet_names)} имён листов из БД.")

            # Обновляем проводник
            if self.file_explorer:
                 self.file_explorer.load_excel_file(self.project_path, self.sheet_names)

        except Exception as e:
            logger.error(f"Ошибка при загрузке имён листов из БД: {e}", exc_info=True)
        finally:
            if self.db_storage:
                self.db_storage.disconnect()

    # --- Возвращаем старые методы, но адаптируем под новую логику ---
    def _on_import_excel(self):
        """Обработчик импорта Excel-файла."""
        if not self.project_db_path:
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
            # Анализируем файл
            logger.info(f"Анализ файла: {file_path}")
            analysis_results = analyze_excel_file(file_path)
            
            if not analysis_results or "sheets" not in analysis_results:
                raise Exception("Анализ не вернул ожидаемые результаты")
            
            # Получаем имена листов
            self.sheet_names = [sheet["name"] for sheet in analysis_results["sheets"]]
            if not self.sheet_names:
                raise Exception("В файле не найдено листов")
            
            # Убедимся, что self.db_storage указывает на правильную БД
            if not self.db_storage or self.db_storage.db_path != self.project_db_path:
                 self.db_storage = ProjectDBStorage(self.project_db_path)

            # Сохраняем результаты анализа в БД (с подключением внутри)
            self._save_analysis_to_db(analysis_results)
            
            # Обновляем состояние
            self.current_excel_file = file_path
            
            # Обновляем UI
            if self.file_explorer:
                self.file_explorer.load_excel_file(file_path, self.sheet_names)
            # self.action_export_excel.setEnabled(True) # Уже включена при открытии/создании проекта
            # self.action_close_file.setEnabled(True) # Уже включена при открытии/создании проекта
            
            # Загружаем первый лист
            self._load_first_sheet()

            logger.info(f"Файл успешно проанализирован: {file_path}, БД: {self.project_db_path}")
            QMessageBox.information(self, "Успех", "Файл успешно проанализирован!")

        except Exception as e:
            logger.error(f"Ошибка при импорте/анализе файла: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось проанализировать файл:\n{e}")
            # _cleanup_resources не вызываем, т.к. проект может остаться валидным

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
    
    def _load_first_sheet(self):
        """Загружает первый лист проекта, если он существует."""
        if self.sheet_names:
            first_sheet_name = self.sheet_names[0]
            logger.info(f"Автоматическая загрузка первого листа: {first_sheet_name}")
            self._on_sheet_selected(first_sheet_name)
        else:
            logger.warning("Нет доступных листов для автоматической загрузки.")

    @Slot(str)
    def _on_sheet_selected(self, sheet_name: str):
        """Обработчик выбора листа."""
        logger.debug(f"Выбран лист: {sheet_name}")
        print(f"[DEBUG MainWin] _on_sheet_selected вызван для листа '{sheet_name}'") # <-- Добавим print
        if not self.project_db_path:
            logger.warning("Нет активной БД")
            print(f"[DEBUG MainWin] _on_sheet_selected: self.project_db_path = {self.project_db_path}") # <-- Добавим print
            return

        print(f"[DEBUG MainWin] _on_sheet_selected: self.sheet_editor = {self.sheet_editor is not None}, self.stacked_widget = {self.stacked_widget is not None}") # <-- Добавим print
        if self.sheet_editor and self.stacked_widget:
            print(f"[DEBUG MainWin] _on_sheet_selected: Переключаюсь на sheet_editor и вызываю load_sheet для '{sheet_name}' с БД {self.project_db_path}") # <-- Добавим print
            self.stacked_widget.setCurrentWidget(self.sheet_editor)
            self.sheet_editor.load_sheet(self.project_db_path, sheet_name)
        else:
            logger.warning("sheet_editor или stacked_widget не инициализированы")
    
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
        # Закрываем соединение с БД
        if self.db_storage:
            self.db_storage.disconnect()
            self.db_storage = None
        
        # Сбрасываем состояние
        self.current_excel_file = None
        self.project_path = None
        self.project_db_path = None
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