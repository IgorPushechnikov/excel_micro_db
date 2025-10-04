# backend/constructor/widgets/new_gui/main_window.py
"""
Главное окно приложения Excel Micro DB GUI.
Управляет проектами, листами и отображает SheetEditorWidget.
"""

import logging
from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QStatusBar, QToolBar, QComboBox, QFileDialog,
    QMessageBox, QProgressDialog, QWidget, QStackedWidget
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QIcon

from backend.core.app_controller import create_app_controller
from .sheet_editor_widget import SheetEditorWidget
from backend.utils.logger import get_logger

logger = get_logger(__name__)

# --- Вспомогательный класс для выполнения анализа в отдельном потоке ---
class AnalysisWorker(QThread):
    """
    Рабочий поток для выполнения анализа Excel-файла.
    Выполняет анализ в отдельном потоке, не взаимодействуя напрямую с AppController.
    """
    # finished = Signal(bool)  # Старый сигнал
    finished = Signal(object, bool)  # Новый сигнал: (результат_анализа, успех/ошибка)

    def __init__(self, excel_file_path, options=None):
        super().__init__()
        self.excel_file_path = excel_file_path
        self.options = options or {} # options пока не используются в анализе, но могут пригодиться

    def run(self):
        """
        Запускает анализ в отдельном потоке.
        """
        try:
            logger.info(f"Начало анализа файла {self.excel_file_path} в потоке {id(QThread.currentThread())}")
            # Импортируем анализатор
            from backend.analyzer.logic_documentation import analyze_excel_file

            # Вызываем анализатор напрямую
            # NOTE: analyze_excel_file принимает только file_path, опции обрабатываются внутри или не поддерживаются напрямую.
            # Для выборочной загрузки нужно будет модифицировать сам analyze_excel_file или AppController.
            # Пока передаём только путь.
            analysis_results = analyze_excel_file(self.excel_file_path) # <-- Убран self.options
            logger.info(f"Анализ файла {self.excel_file_path} завершён в потоке {id(QThread.currentThread())}.")

            # Отправляем результат анализа и флаг успеха
            self.finished.emit(analysis_results, True)
        except Exception as e:
            logger.error(f"Ошибка в потоке анализа для файла {self.excel_file_path}: {e}", exc_info=True)
            # Отправляем None и флаг ошибки
            self.finished.emit(None, False)

# --- Основное окно ---
class MainWindow(QMainWindow):
    """
    Главное окно GUI приложения.
    """

    def __init__(self):
        super().__init__()
        self.app_controller = create_app_controller()
        self.current_project_path = None
        self.current_sheet_name = None
        self.sheet_editor_widget = None
        self.analysis_worker = None
        self.progress_dialog = None
        # NOTE: Добавим атрибут для хранения пути к файлу, который нужно импортировать
        # Он будет установлен в _on_import_triggered и использован в _on_analysis_finished
        self.excel_file_path_to_import = None
        self.import_options = None

        self.setWindowTitle("Excel Micro DB GUI")
        self.setGeometry(100, 100, 1200, 800)

        self._setup_ui()
        self._setup_connections()

    def _setup_ui(self):
        """
        Создаёт элементы интерфейса.
        """
        # Меню
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)

        file_menu = menu_bar.addMenu("Файл")
        import_menu = menu_bar.addMenu("Импорт")
        export_menu = menu_bar.addMenu("Экспорт")

        # Действия
        self.new_project_action = QAction("Новый проект", self)
        self.open_project_action = QAction("Открыть проект", self)
        self.save_project_action = QAction("Сохранить проект", self)
        self.import_action = QAction("Импорт", self) # Основное действие, которое будет вызывать выпадающий список
        self.export_action = QAction("Экспорт", self)

        # Добавление действий в меню
        file_menu.addAction(self.new_project_action)
        file_menu.addAction(self.open_project_action)
        file_menu.addAction(self.save_project_action)
        import_menu.addAction(self.import_action)
        export_menu.addAction(self.export_action)

        # Панель инструментов
        tool_bar = QToolBar("Основная", self)
        self.addToolBar(tool_bar)

        # Комбобокс для выбора листа
        self.sheet_combo_box = QComboBox(self)
        self.sheet_combo_box.addItem("Выберите лист...")
        tool_bar.addWidget(self.sheet_combo_box)

        # Выпадающий список для импорта
        self.import_combo_box = QComboBox(self)
        self.import_combo_box.addItems(["Все данные", "Последовательный", "Выборочный"])
        # Добавляем его как кнопку с выпадающим списком
        import_button = tool_bar.addWidget(self.import_combo_box)
        import_button.setText("Импорт")

        # Статусная строка
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов.")

        # Центральный виджет (для SheetEditorWidget)
        self.stacked_widget = QStackedWidget(self)
        placeholder_widget = QWidget(self)
        # Можно добавить QLabel с приветствием или инструкцией
        self.stacked_widget.addWidget(placeholder_widget)
        self.setCentralWidget(self.stacked_widget)

    def _setup_connections(self):
        """
        Подключает сигналы к слотам.
        """
        self.new_project_action.triggered.connect(self._on_new_project_triggered)
        self.open_project_action.triggered.connect(self._on_open_project_triggered)
        self.save_project_action.triggered.connect(self._on_save_project_triggered)
        self.import_action.triggered.connect(self._on_import_triggered)
        self.export_action.triggered.connect(self._on_export_triggered)
        self.sheet_combo_box.currentTextChanged.connect(self._on_sheet_changed)

    def _on_new_project_triggered(self):
        """
        Обработчик создания нового проекта.
        """
        project_path_str, ok = QFileDialog.getSaveFileName(
            self, "Новый проект", "", "Папки проектов ();;Все файлы (*)", options=QFileDialog.Option.DontUseNativeDialog
        )
        if ok and project_path_str:
            project_path = Path(project_path_str)
            if project_path.suffix: # Если пользователь ввёл имя файла, а не папку
                project_path = project_path.parent / project_path.stem # Берём только папку
            try:
                # Создаём проект через AppController
                success = self.app_controller.create_project(str(project_path))
                if success:
                    logger.info(f"Новый проект создан: {project_path}")
                    self.current_project_path = project_path
                    # После создания проекта его нужно загрузить
                    self.app_controller.load_project(str(project_path))
                    self._update_sheet_list()
                    self.status_bar.showMessage(f"Проект создан: {project_path}")
                else:
                    logger.error(f"Не удалось создать проект: {project_path}")
                    QMessageBox.critical(self, "Ошибка", f"Не удалось создать проект: {project_path}")
            except Exception as e:
                logger.error(f"Ошибка при создании проекта: {e}", exc_info=True)
                QMessageBox.critical(self, "Ошибка", f"Ошибка при создании проекта: {e}")

    def _on_open_project_triggered(self):
        """
        Обработчик открытия существующего проекта.
        """
        # --- ИСПРАВЛЕНО: getExistingDirectory возвращает только строку ---
        project_path_str = QFileDialog.getExistingDirectory(
            self, "Открыть проект", "", options=QFileDialog.Option.DontUseNativeDialog
        )
        # Проверяем, пустой ли путь (означает отмену)
        if not project_path_str:
            return
        # --- КОНЕЦ ИСПРАВЛЕНИЯ ---
        project_path = Path(project_path_str)
        try:
            # Загружаем проект через AppController
            success = self.app_controller.load_project(str(project_path))
            if success:
                logger.info(f"Проект загружен: {project_path}")
                self.current_project_path = project_path
                self._update_sheet_list()
                self.status_bar.showMessage(f"Проект загружен: {project_path}")
            else:
                logger.error(f"Не удалось загрузить проект: {project_path}")
                QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить проект: {project_path}")
        except Exception as e:
            logger.error(f"Ошибка при загрузке проекта: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке проекта: {e}")

    def _on_save_project_triggered(self):
        """
        Обработчик сохранения проекта.
        """
        # AppController управляет сохранением через методы типа update_cell_value
        # и, возможно, через экспорт. Здесь можно реализовать сохранение изменений в БД,
        # если есть отложенные операции, или просто вызвать экспорт.
        # Пока просто покажем сообщение.
        if self.current_project_path:
            self.status_bar.showMessage(f"Проект сохранён: {self.current_project_path}")
            logger.info(f"Проект сохранён: {self.current_project_path}")
        else:
            QMessageBox.information(self, "Сохранение", "Нет активного проекта для сохранения.")

    def _on_import_triggered(self):
        """
        Обработчик импорта.
        """
        # Открываем диалог выбора файла Excel
        file_path_str, ok = QFileDialog.getOpenFileName(
            self, "Импорт Excel", "", "Excel Files (*.xlsx *.xls);;Все файлы (*)", options=QFileDialog.Option.DontUseNativeDialog
        )
        if ok and file_path_str:
            excel_file_path = Path(file_path_str)
            if not excel_file_path.exists():
                QMessageBox.critical(self, "Ошибка", f"Файл не найден: {excel_file_path}")
                return

            # Получаем выбранный тип импорта из выпадающего списка
            import_type = self.import_combo_box.currentText()
            logger.info(f"Выбран тип импорта: {import_type} для файла {excel_file_path}")

            # Показываем диалог прогресса
            self.progress_dialog = QProgressDialog("Анализ файла...", "Отмена", 0, 100, self)
            self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress_dialog.setMinimumDuration(0) # Показываем сразу
            self.progress_dialog.setValue(0)

            # --- ИСПРАВЛЕНО: Передаём excel_file_path и options в AnalysisWorker ---
            # Запоминаем путь и опции для использования в _on_analysis_finished
            self.excel_file_path_to_import = str(excel_file_path)
            self.import_options = {'import_type': import_type}
            self.analysis_worker = AnalysisWorker(
                str(excel_file_path), # Передаём путь к файлу
                options={'import_type': import_type} # Передаём тип импорта как опцию (пока не используется в анализе)
            )
            # --- КОНЕЦ ИСПРАВЛЕНИЯ ---
            self.analysis_worker.finished.connect(self._on_analysis_finished)
            self.analysis_worker.start()

    def _on_analysis_progress(self, value, message):
        """
        Обработчик прогресса анализа.
        """
        if self.progress_dialog:
            self.progress_dialog.setValue(value)
            self.progress_dialog.setLabelText(message)
            # Проверяем, не нажата ли кнопка отмены
            if self.progress_dialog.wasCanceled():
                # AppController должен поддерживать отмену, если возможно
                # Пока просто останавливаем поток (не идеально)
                # --- ИСПРАВЛЕНО: Проверка на None перед terminate и wait ---
                if self.analysis_worker:
                    self.analysis_worker.terminate()
                    self.analysis_worker.wait()
                # --- КОНЕЦ ИСПРАВЛЕНИЯ ---
                self.progress_dialog = None
                self.analysis_worker = None
                logger.info("Анализ отменён пользователем.")
                self.status_bar.showMessage("Анализ отменён.")
                return

    def _on_analysis_finished(self, analysis_results, success):
        """
        Обработчик завершения анализа.
        """
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None

        if self.analysis_worker:
            self.analysis_worker.wait() # Убедиться, что поток завершён
            self.analysis_worker = None

        if success and analysis_results is not None:
            logger.info("Анализ файла успешно завершён в отдельном потоке. Начинаю сохранение в БД через AppController в основном потоке.")
            # --- НОВОЕ: Вызываем AppController в основном потоке ---
            # NOTE: Здесь мы снова вызываем analyze_excel_file, который сам анализирует файл.
            # Это неэффективно. Лучше было бы модифицировать AppController,
            # чтобы он мог принимать уже проанализированные данные (analysis_results).
            # Но для простоты и соответствия текущей структуре AppController,
            # мы вызываем его снова. Главное, что теперь это происходит в основном потоке.
            # TODO: Рассмотреть передачу analysis_results в AppController.
            try:
                # Вызов AppController.analyze_excel_file в основном потоке
                # Используем путь и опции, сохранённые в _on_import_triggered
                app_controller_success = self.app_controller.analyze_excel_file(self.excel_file_path_to_import, self.import_options)
                if app_controller_success:
                    logger.info("Сохранение результатов анализа в БД через AppController успешно завершено.")
                    self.status_bar.showMessage("Импорт завершён успешно.")
                    # Обновляем список листов, так как могли появиться новые
                    self._update_sheet_list()
                else:
                    logger.error("AppController не смог сохранить результаты анализа в БД.")
                    self.status_bar.showMessage("Ошибка сохранения в БД после анализа.")
                    QMessageBox.critical(self, "Ошибка", "AppController не смог сохранить результаты анализа в БД.")
            except Exception as e_app:
                logger.error(f"Ошибка при вызове AppController.analyze_excel_file: {e_app}", exc_info=True)
                self.status_bar.showMessage("Ошибка вызова AppController.")
                QMessageBox.critical(self, "Ошибка", f"Ошибка при вызове AppController: {e_app}")
            # --- КОНЕЦ НОВОГО ---
        else:
            logger.error("Анализ файла завершился с ошибкой в отдельном потоке или результат пуст.")
            self.status_bar.showMessage("Ошибка импорта.")
            QMessageBox.critical(self, "Ошибка", "Ошибка при импорте файла.")

    def _on_export_triggered(self):
        """
        Обработчик экспорта.
        """
        if not self.current_project_path:
            QMessageBox.information(self, "Экспорт", "Нет активного проекта для экспорта.")
            return

        output_path_str, ok = QFileDialog.getSaveFileName(
            self, "Экспорт Excel", "", "Excel Files (*.xlsx);;Все файлы (*)", options=QFileDialog.Option.DontUseNativeDialog
        )
        if ok and output_path_str:
            output_path = Path(output_path_str)
            try:
                # Вызываем экспорт через AppController
                success = self.app_controller.export_results(export_type='excel', output_path=str(output_path))
                if success:
                    logger.info(f"Проект экспортирован в: {output_path}")
                    self.status_bar.showMessage(f"Проект экспортирован: {output_path}")
                else:
                    logger.error(f"Не удалось экспортировать проект в: {output_path}")
                    QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать проект: {output_path}")
            except Exception as e:
                logger.error(f"Ошибка при экспорте проекта: {e}", exc_info=True)
                QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте проекта: {e}")

    def _update_sheet_list(self):
        """
        Обновляет список листов в комбобоксе.
        """
        if not self.app_controller.is_project_loaded:
            self.sheet_combo_box.clear()
            self.sheet_combo_box.addItem("Нет проекта")
            return

        try:
            sheet_names = self.app_controller.get_sheet_names()
            logger.debug(f"Получены имена листов: {sheet_names}")
            self.sheet_combo_box.clear()
            if sheet_names:
                self.sheet_combo_box.addItems(sheet_names)
            else:
                self.sheet_combo_box.addItem("Нет листов")
        except Exception as e:
            logger.error(f"Ошибка при получении списка листов: {e}", exc_info=True)
            self.sheet_combo_box.clear()
            self.sheet_combo_box.addItem("Ошибка загрузки")

    def _on_sheet_changed(self, sheet_name: str):
        """
        Обработчик смены активного листа.
        """
        if not sheet_name or sheet_name in ["Выберите лист...", "Нет проекта", "Нет листов", "Ошибка загрузки"]:
            return

        logger.info(f"Смена активного листа на: {sheet_name}")
        self.current_sheet_name = sheet_name

        # Удаляем предыдущий SheetEditorWidget, если он был
        if self.sheet_editor_widget:
            self.stacked_widget.removeWidget(self.sheet_editor_widget)
            self.sheet_editor_widget.deleteLater() # Удаляем из Qt
            self.sheet_editor_widget = None

        # Создаём новый SheetEditorWidget для выбранного листа
        self.sheet_editor_widget = SheetEditorWidget(self.app_controller, sheet_name, self)
        self.stacked_widget.addWidget(self.sheet_editor_widget)
        self.stacked_widget.setCurrentWidget(self.sheet_editor_widget)

        self.status_bar.showMessage(f"Активный лист: {sheet_name}")
