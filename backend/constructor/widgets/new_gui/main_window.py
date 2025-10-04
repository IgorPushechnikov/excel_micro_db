# backend/constructor/widgets/new_gui/main_window.py
"""
Модуль для главного окна приложения.
Содержит класс MainWindow, который является основным графическим интерфейсом.
"""

import os
import logging
import threading
from typing import List, Optional
# --- Импорты PyQt5 ---
from PyQt5.QtWidgets import (
    QMainWindow, QFileDialog, QVBoxLayout, QHBoxLayout, QWidget,
    QToolBar, QAction, QStatusBar, QLabel, QStackedWidget,
    QTreeWidget, QTreeWidgetItem # <-- ЗАМЕНА: QListWidget -> QTreeWidget/QTreeWidgetItem
)
from PyQt5.QtCore import Qt, pyqtSignal, QObject
from PyQt5.QtGui import QIcon

# --- Импорты проекта ---
# ИСПРАВЛЕНО: Корректные пути импортов
from backend.core.app_controller import AppController
from backend.constructor.widgets.new_gui.qt_model_adapter import QtModelAdapter
from backend.constructor.widgets.new_gui.sheet_editor_widget import SheetEditorWidget

logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """
    Главное окно приложения.
    
    Этот класс управляет всем пользовательским интерфейсом:
    - Создание меню и панелей инструментов.
    - Управление состоянием проекта (открытие, сохранение).
    - Отображение редактора листов.
    """

    # Сигнал для асинхронного обновления UI после анализа
    analysis_finished_signal = pyqtSignal(list) # Испускается, когда анализ завершен

    def __init__(self, app_controller: AppController):
        super().__init__()
        self.app_controller = app_controller
        self.setWindowTitle("Excel Micro-DB")
        self.setGeometry(100, 100, 800, 600)

        # --- Атрибуты ---
        self.current_project_path: Optional[str] = None
        self._current_sheet_name: Optional[str] = None
        self.model_adapters: Dict[str, QtModelAdapter] = {} # Хранит адаптеры моделей для каждого листа

        # --- Виджеты ---
        # Центральный виджет и layout
        central_widget = QWidget()
        main_layout = QHBoxLayout() # Основной layout
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Панель для выбора листов (слева)
        sheet_panel = QWidget()
        sheet_panel.setMaximumWidth(200) # Максимальная ширина панели
        sheet_layout = QVBoxLayout()
        sheet_panel.setLayout(sheet_layout)

        # Новый виджет для списка листов
        self.sheet_tree_widget = QTreeWidget() # <-- ИСПОЛЬЗУЕТСЯ QTreeWidget
        self.sheet_tree_widget.setHeaderLabel("Листы") # <-- Установка заголовка
        self.sheet_tree_widget.setHeaderHidden(True) # <-- Скрытие заголовков столбцов
        self.sheet_tree_widget.setColumnCount(1) # <-- Один столбец
        # Подключение сигнала выбора элемента
        self.sheet_tree_widget.itemSelectionChanged.connect(self._on_sheet_selected)

        sheet_layout.addWidget(self.sheet_tree_widget)

        # Центральный виджет (редактор листов)
        self.stacked_widget = QStackedWidget()

        # Добавляем панели листов и редактора в splitter
        splitter = QSplitter(Qt.Horizontal) # Горизонтальный сплиттер
        splitter.addWidget(sheet_panel)
        splitter.addWidget(self.stacked_widget)
        splitter.setSizes([200, 600]) # Инициализация размеров
        main_layout.addWidget(splitter)

        # --- Меню и панель инструментов ---
        self._setup_menu_bar()
        self._setup_tool_bar()

        # --- Статус-бар ---
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.project_status_label = QLabel("Проект не загружен")
        self.statusBar.addPermanentWidget(self.project_status_label)

        # --- Сигналы ---
        self.analysis_finished_signal.connect(self._on_analysis_finished_in_main_thread)

        # --- Инициализация ---
        if not self.app_controller.initialize():
            logger.error("Не удалось инициализировать AppController")
            self.close()


    def _setup_menu_bar(self):
        """Настраивает меню бар."""
        menu_bar = self.menuBar()

        # Меню Файл
        file_menu = menu_bar.addMenu('Файл')

        # Действия
        new_action = QAction(QIcon(), 'Новый проект', self)
        new_action.triggered.connect(self._on_new_project)
        file_menu.addAction(new_action)

        open_action = QAction(QIcon(), 'Открыть проект', self)
        open_action.triggered.connect(self._on_open_project)
        file_menu.addAction(open_action)

        save_action = QAction(QIcon(), 'Сохранить проект', self)
        save_action.triggered.connect(self._on_save_project)
        file_menu.addAction(save_action)

        export_action = QAction(QIcon(), 'Экспорт проекта', self)
        export_action.triggered.connect(self._on_export_project)
        file_menu.addAction(export_action)

        file_menu.addSeparator()
        exit_action = QAction(QIcon(), 'Выход', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Меню Проект
        project_menu = menu_bar.addMenu('Проект')
        import_action = QAction(QIcon(), 'Импорт Excel', self)
        import_action.triggered.connect(self._on_import_excel)
        project_menu.addAction(import_action)

        # Меню Анализ
        analyze_menu = menu_bar.addMenu('Анализ')
        run_analyze_action = QAction(QIcon(), 'Запустить анализ', self)
        run_analyze_action.triggered.connect(self._on_run_analysis)
        analyze_menu.addAction(run_analyze_action)


    def _setup_tool_bar(self):
        """Настраивает панель инструментов."""
        tool_bar = QToolBar("Основная панель", self)
        self.addToolBar(tool_bar)

        # Действия
        new_action = QAction(QIcon(), 'Новый', self)
        new_action.triggered.connect(self._on_new_project)
        tool_bar.addAction(new_action)

        open_action = QAction(QIcon(), 'Открыть', self)
        open_action.triggered.connect(self._on_open_project)
        tool_bar.addAction(open_action)

        save_action = QAction(QIcon(), 'Сохранить', self)
        save_action.triggered.connect(self._on_save_project)
        tool_bar.addAction(save_action)

        tool_bar.addSeparator()

        import_action = QAction(QIcon(), 'Импорт', self)
        import_action.triggered.connect(self._on_import_excel)
        tool_bar.addAction(import_action)

        # Не добавляем комбо-бокс на tool_bar
        # tool_bar.addWidget(self.sheet_combo_box) # <-- ЭТОЙ СТРОКИ БОЛЬШЕ НЕТ

    def _update_sheet_list(self, sheet_names: List[str]):
        """Обновляет список листов в выпадающем меню.
        Args:
            sheet_names (List[str]): Список имен листов.
        """
        logger.debug(f"Получены имена листов: {sheet_names}")

        # --- НОВОЕ: Очистка и заполнение QTreeWidget ---
        self.sheet_tree_widget.clear()
        for name in sheet_names:
            item = QTreeWidgetItem([name]) # Создаем элемент с текстом
            self.sheet_tree_widget.addTopLevelItem(item)
        # ---------------------------------------------------

    def _on_sheet_changed(self, sheet_name: str):
        """Обрабатывает изменение активного листа.
        Args:
            sheet_name (str): Имя выбранного листа.
        """
        logger.info(f"Смена активного листа на: {sheet_name}")
        self._current_sheet_name = sheet_name
        # Логика переключения виджетов
        editor_widget = self.stacked_widget.findChild(SheetEditorWidget, sheet_name)
        if editor_widget:
            self.stacked_widget.setCurrentWidget(editor_widget)
            # Обновляем статус-бар
            self.project_status_label.setText(f"Текущий проект: {os.path.basename(self.current_project_path)} | Лист: {sheet_name}")
        else:
            logger.warning(f"Редактор для листа '{sheet_name}' не найден.")

    def _on_sheet_selected(self):
        """Обрабатывает выбор листа из tree widget.
        """
        selected_items = self.sheet_tree_widget.selectedItems()
        if selected_items:
            selected_sheet_name = selected_items[0].text(0) # Текст первого столбца
            self._on_sheet_changed(selected_sheet_name)

    def _create_or_update_sheet_editor(self, sheet_name: str):
        """Создает или обновляет редактор для листа.
        Args:
            sheet_name (str): Имя листа.
        """
        # Ищем существующий редактор
        existing_editor = self.stacked_widget.findChild(SheetEditorWidget, sheet_name)
        if existing_editor:
            # Если существует, просто обновляем его модель
            model_adapter = self.model_adapters.get(sheet_name)
            if model_adapter:
                existing_editor.set_model(model_adapter)
                logger.info(f"Обновлен редактор для существующего листа: {sheet_name}")
        else:
            # Создаем новый редактор
            model_adapter = QtModelAdapter(sheet_name, self.app_controller)
            self.model_adapters[sheet_name] = model_adapter
            editor_widget = SheetEditorWidget(model_adapter, parent=self)
            editor_widget.setObjectName(sheet_name) # Устанавливаем имя объекта
            self.stacked_widget.addWidget(editor_widget)
            logger.info(f"Создан редактор для листа: {sheet_name}")

        # Принудительно переключаемся на этот лист
        self._on_sheet_changed(sheet_name)

    def _load_sheets_from_controller(self):
        """Загружает список листов из AppController и обновляет UI."""
        try:
            sheet_names = self.app_controller.get_sheet_names()
            self._update_sheet_list(sheet_names)
            if sheet_names:
                first_sheet = sheet_names[0]
                self._create_or_update_sheet_editor(first_sheet)
                self._on_sheet_changed(first_sheet)
            else:
                logger.info("Нет листов для загрузки.")
        except Exception as e:
            logger.error(f"Ошибка при загрузке листов из контроллера: {e}", exc_info=True)

    # --- Слоты для действий меню ---
    def _on_new_project(self):
        """Обработка создания нового проекта."""
        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(self, "Выбор директории для нового проекта")
        if folder_path:
            project_name = os.path.basename(folder_path)
            success = self.app_controller.create_new_project(project_name)
            if success:
                self.current_project_path = folder_path
                self._load_sheets_from_controller()
                self.project_status_label.setText(f"Новый проект создан: {project_name}")
                logger.info(f"Новый проект создан: {folder_path}")
            else:
                logger.error("Не удалось создать новый проект.")

    def _on_open_project(self):
        """Обработка открытия существующего проекта."""
        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(self, "Выбор директории проекта")
        if folder_path:
            success = self.app_controller.load_project(folder_path)
            if success:
                self.current_project_path = folder_path
                self._load_sheets_from_controller()
                self.project_status_label.setText(f"Проект загружен: {os.path.basename(folder_path)}")
                logger.info(f"Проект загружен: {folder_path}")
            else:
                logger.error("Не удалось загрузить проект.")

    def _on_save_project(self):
        """Обработка сохранения проекта."""
        if self.app_controller.is_project_loaded and self.current_project_path:
            # Предполагаем, что AppController сам знает, куда сохраняться
            # или мы передаем ему путь
            # Для MVP вызываем напрямую
            # TODO: Реализовать настоящий механизм сохранения в ProjectManager
            self.app_controller.save_project() # <-- Пока это заглушка
            self.project_status_label.setText(f"Проект сохранён: {os.path.basename(self.current_project_path)}")
            logger.info(f"Проект сохранён: {self.current_project_path}")
        else:
            logger.warning("Попытка сохранения без загруженного проекта.")

    def _on_export_project(self):
        """Обработка экспорта проекта в Excel."""
        if not self.app_controller.is_project_loaded:
            logger.warning("Попытка экспорта без загруженного проекта.")
            return

        options = QFileDialog.Options()
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Экспорт проекта",
            "test.xlsx", # <-- Заменил на более понятное имя по умолчанию
            "Excel Files (*.xlsx);;All Files (*)",
            options=options
        )
        if output_path:
            # Используем универсальный метод экспортa
            success = self.app_controller.export_results('excel', output_path)
            if success:
                logger.info(f"Проект экспортирован в: {output_path}")
                self.statusBar().showMessage(f"Экспорт успешен: {output_path}", 5000)
            else:
                logger.error(f"Не удалось экспортировать проект в: {output_path}")
                self.statusBar().showMessage(f"Ошибка экспорта: {output_path}", 5000)

    def _on_import_excel(self):
        """Обработка импорта Excel-файла."""
        if not self.app_controller.is_project_loaded:
            logger.warning("Попытка импорта без загруженного проекта.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выбор Excel-файла для импорта",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)",
            options=options
        )
        if file_path:
            # Выбор типа импорта
            logger.info(f"Выбран тип импорта: Все данные для файла {file_path}")
            # --- Запуск анализа в отдельном потоке ---
            self._start_analysis_in_thread(file_path, {'mode': 'all'})

    def _start_analysis_in_thread(self, excel_file_path: str, options: dict):
        """Запускает анализ в фоновом потоке.
        Args:
            excel_file_path (str): Путь к Excel-файлу.
            options (dict): Опции анализа.
        """
        thread_id = id(threading.current_thread())
        logger.info(f"Начало анализа файла {excel_file_path} в потоке {thread_id}")

        def run_analysis():
            try:
                # ИСПРАВЛЕНО: Теперь принимает два аргумента
                from backend.analyzer.logic_documentation import analyze_excel_file
                analysis_results = analyze_excel_file(excel_file_path, options)
                # Эмитируем сигнал для обновления UI в основном потоке
                self.analysis_finished_signal.emit(analysis_results)
            except Exception as e:
                logger.error(f"Ошибка в потоке анализа для файла {excel_file_path}: {e}", exc_info=True)

        # Запускаем анализ в отдельном потоке
        thread = threading.Thread(target=run_analysis)
        thread.daemon = True
        thread.start()

    def _on_analysis_finished_in_main_thread(self, analysis_results: list):
        """Обработка результатов анализа в основном потоке.
        Args:
            analysis_results (list): Результаты анализа.
        """
        if not analysis_results:
            logger.error("Анализ файла завершился с ошибкой в отдельном потоке или результат пуст.")
            return

        # Сохраняем результаты в БД через AppController
        logger.info("Анализ файла успешно завершён в потоке. Начинаю сохранение в БД через AppController в основном потоке.")
        # --- НОВОЕ: Сохранение результатов через AppController ---
        success = self.app_controller.analyze_excel_file(
            analysis_results["file_path"], 
            options={"mode": "all"} # Можно использовать опции из results
        )
        # -----------------------------------------------------
        if success:
            logger.info("Сохранение результатов анализа в БД через AppController успешно завершено.")
            # Обновляем список листов
            self._load_sheets_from_controller()
        else:
            logger.error("Не удалось сохранить результаты анализа в БД.")

    def closeEvent(self, event):
        """Обработка закрытия окна."""
        # Здесь можно добавить проверку на несохраненные изменения
        logger.info(f"GUI приложение завершено с кодом: {event.type()}")
        self.app_controller.shutdown()
        event.accept()
