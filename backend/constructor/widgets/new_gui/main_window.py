# backend/constructor/widgets/new_gui/main_window.py
"""
Главное окно приложения Excel Micro DB GUI.
Управляет проектами, листами и отображает SheetEditorWidget.
"""

import logging
from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QStatusBar, QToolBar, QFileDialog,
    QMessageBox, QProgressDialog, QWidget, QStackedWidget, QHBoxLayout, QSplitter,
    QCheckBox, QPushButton, QMenu, QAction, QActionGroup # <-- Новые импорты
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QIcon

from backend.core.app_controller import create_app_controller
from .sheet_editor_widget import SheetEditorWidget
from .sheet_explorer_widget import SheetExplorerWidget # <-- Новый импорт
from backend.utils.logger import get_logger

logger = get_logger(__name__)

# --- НОВОЕ: Вспомогательный класс для выполнения импорта в отдельном потоке ---
class ImportWorker(QThread):
    """
    Рабочий поток для выполнения импорта данных через AppController.
    """
    finished = Signal(bool, str)  # (успех/ошибка, сообщение)
    progress = Signal(int, str)   # (значение, сообщение) - если AppController будет передавать прогресс

    def __init__(self, app_controller, file_path, import_type, import_mode):
        super().__init__()
        self.app_controller = app_controller
        self.file_path = file_path
        self.import_type = import_type
        self.import_mode = import_mode

    def run(self):
        """
        Запускает импорт в отдельном потоке.
        """
        try:
            logger.info(f"Начало импорта (тип: {self.import_type}, режим: {self.import_mode}) для файла {self.file_path} в потоке {id(QThread.currentThread())}")

            # Определяем метод AppController на основе типа и режима
            method_name = f"import_{self.import_type}_from_excel"
            if self.import_mode != 'all' and self.import_mode != 'fast':
                 method_name += f"_{self.import_mode}"

            method = getattr(self.app_controller, method_name, None)
            if method is None:
                raise AttributeError(f"AppController не имеет метода {method_name}")

            # Вызываем метод
            success = method(self.file_path)

            logger.info(f"Импорт (тип: {self.import_type}, режим: {self.import_mode}) для файла {self.file_path} завершён в потоке {id(QThread.currentThread())}.")

            # Отправляем результат
            self.finished.emit(success, f"Импорт ({self.import_type}, {self.import_mode}) {'успешен' if success else 'неудачен'}.")

        except Exception as e:
            logger.error(f"Ошибка в потоке импорта для файла {self.file_path} (тип: {self.import_type}, режим: {self.import_mode}): {e}", exc_info=True)
            # Отправляем ошибку
            self.finished.emit(False, f"Ошибка импорта: {e}")
# --- КОНЕЦ НОВОГО ---

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
        # --- УДАЛЕНО: analysis_worker, excel_file_path_to_import, import_options ---
        # Старый AnalysisWorker и связанные атрибуты больше не нужны
        # self.analysis_worker = None
        # self.excel_file_path_to_import = None
        # self.import_options = None
        # --- КОНЕЦ УДАЛЕНИЯ ---
        self.import_worker = None # <-- Новый атрибут для нового потока импорта
        self.progress_dialog = None

        # --- НОВОЕ: Атрибуты для хранения выбранного типа и режима импорта ---
        self.selected_import_type = 'all_data'  # 'all_data', 'styles', 'charts', 'formulas', 'raw_data_pandas'
        self.selected_import_mode = 'all'       # 'all', 'selective', 'chunks', 'fast'
        # --- КОНЕЦ НОВОГО ---

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

        # --- НОВОЕ: Чекбокс для управления логированием ---
        self.logging_checkbox = QCheckBox("Логирование", self)
        self.logging_checkbox.setChecked(True)  # По умолчанию включено
        tool_bar.addWidget(self.logging_checkbox)
        # --- КОНЕЦ НОВОГО ---

        # --- УДАЛЕНО: Старый import_combo_box ---
        # self.import_combo_box = QComboBox(self)
        # self.import_combo_box.addItems(["Все данные", "Последовательный", "Выборочный"])
        # import_button = tool_bar.addWidget(self.import_combo_box)
        # import_button.setText("Импорт")
        # --- КОНЕЦ УДАЛЕНИЯ ---

        # --- НОВОЕ: Кнопка с меню для выбора типа и режима импорта ---
        self.import_type_button = QPushButton("Тип импорта", self)
        self.import_type_menu = QMenu(self.import_type_button)

        # Подменю для типа данных
        self.import_data_type_menu = self.import_type_menu.addMenu("Тип данных")
        self.import_data_type_action_group = QActionGroup(self)
        self.import_data_type_action_group.setExclusive(True)

        self.all_data_action = QAction("Все данные", self)
        self.all_data_action.setCheckable(True)
        self.all_data_action.setChecked(True) # По умолчанию
        self.all_data_action.triggered.connect(lambda: self._on_import_type_selected('all_data'))
        self.import_data_type_action_group.addAction(self.all_data_action)
        self.import_data_type_menu.addAction(self.all_data_action)

        self.raw_data_action = QAction("Сырые данные", self)
        self.raw_data_action.setCheckable(True)
        self.raw_data_action.triggered.connect(lambda: self._on_import_type_selected('raw_data'))
        self.import_data_type_action_group.addAction(self.raw_data_action)
        self.import_data_type_menu.addAction(self.raw_data_action)

        self.styles_action = QAction("Стили", self)
        self.styles_action.setCheckable(True)
        self.styles_action.triggered.connect(lambda: self._on_import_type_selected('styles'))
        self.import_data_type_action_group.addAction(self.styles_action)
        self.import_data_type_menu.addAction(self.styles_action)

        self.charts_action = QAction("Диаграммы", self)
        self.charts_action.setCheckable(True)
        self.charts_action.triggered.connect(lambda: self._on_import_type_selected('charts'))
        self.import_data_type_action_group.addAction(self.charts_action)
        self.import_data_type_menu.addAction(self.charts_action)

        self.formulas_action = QAction("Формулы", self)
        self.formulas_action.setCheckable(True)
        self.formulas_action.triggered.connect(lambda: self._on_import_type_selected('formulas'))
        self.import_data_type_action_group.addAction(self.formulas_action)
        self.import_data_type_menu.addAction(self.formulas_action)

        self.raw_data_pandas_action = QAction("Сырые данные (pandas)", self)
        self.raw_data_pandas_action.setCheckable(True)
        self.raw_data_pandas_action.triggered.connect(lambda: self._on_import_type_selected('raw_data_pandas'))
        self.import_data_type_action_group.addAction(self.raw_data_pandas_action)
        self.import_data_type_menu.addAction(self.raw_data_pandas_action)

        # Подменю для режима импорта
        self.import_mode_menu = self.import_type_menu.addMenu("Режим импорта")
        self.import_mode_action_group = QActionGroup(self)
        self.import_mode_action_group.setExclusive(True)

        self.import_mode_all_action = QAction("Всё", self)
        self.import_mode_all_action.setCheckable(True)
        self.import_mode_all_action.setChecked(True) # По умолчанию
        self.import_mode_all_action.triggered.connect(lambda: self._on_import_mode_selected('all'))
        self.import_mode_action_group.addAction(self.import_mode_all_action)
        self.import_mode_menu.addAction(self.import_mode_all_action)

        self.import_mode_selective_action = QAction("Выборочно", self)
        self.import_mode_selective_action.setCheckable(True)
        self.import_mode_selective_action.triggered.connect(lambda: self._on_import_mode_selected('selective'))
        self.import_mode_action_group.addAction(self.import_mode_selective_action)
        self.import_mode_menu.addAction(self.import_mode_selective_action)

        self.import_mode_chunks_action = QAction("Частями", self)
        self.import_mode_chunks_action.setCheckable(True)
        self.import_mode_chunks_action.triggered.connect(lambda: self._on_import_mode_selected('chunks'))
        self.import_mode_action_group.addAction(self.import_mode_chunks_action)
        self.import_mode_menu.addAction(self.import_mode_chunks_action)

        self.import_mode_fast_action = QAction("Быстрый (pandas)", self)
        self.import_mode_fast_action.setCheckable(True)
        self.import_mode_fast_action.setEnabled(False) # Доступен только для 'raw_data_pandas'
        self.import_mode_fast_action.triggered.connect(lambda: self._on_import_mode_selected('fast'))
        self.import_mode_action_group.addAction(self.import_mode_fast_action)
        self.import_mode_menu.addAction(self.import_mode_fast_action)

        # Привязываем меню к кнопке
        self.import_type_button.setMenu(self.import_type_menu)

        # Добавляем кнопку на панель инструментов
        tool_bar.addWidget(self.import_type_button)
        # --- КОНЕЦ НОВОГО ---

        # Статусная строка
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов.")

        # --- НОВОЕ: Центральный виджет с разделителем (Splitter) ---
        central_widget = QWidget(self)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Создаем обозреватель листов
        self.sheet_explorer = SheetExplorerWidget(self.app_controller, self)

        # Создаем стек для редакторов листов
        self.stacked_widget = QStackedWidget(self)
        placeholder_widget = QWidget(self)
        self.stacked_widget.addWidget(placeholder_widget)

        # Создаем разделитель
        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.addWidget(self.sheet_explorer)
        splitter.addWidget(self.stacked_widget)
        splitter.setSizes([200, 1000]) # Примерное соотношение ширины

        main_layout.addWidget(splitter)
        self.setCentralWidget(central_widget)
        # --- КОНЕЦ НОВОГО ---

    def _setup_connections(self):
        """
        Подключает сигналы к слотам.
        """
        self.new_project_action.triggered.connect(self._on_new_project_triggered)
        self.open_project_action.triggered.connect(self._on_open_project_triggered)
        self.save_project_action.triggered.connect(self._on_save_project_triggered)
        self.import_action.triggered.connect(self._on_import_triggered)
        self.export_action.triggered.connect(self._on_export_triggered)

        # --- НОВОЕ: Подключение чекбокса логирования ---
        self.logging_checkbox.stateChanged.connect(self._on_logging_toggled)
        # --- КОНЕЦ НОВОГО ---

        # --- НОВОЕ: Подключение сигналов обозревателя ---
        self.sheet_explorer.sheet_selected.connect(self._on_sheet_selected)
        self.sheet_explorer.sheet_renamed.connect(self._on_sheet_renamed)
        # --- КОНЕЦ НОВОГО ---

        # --- НОВОЕ: Подключение сигналов выбора типа и режима импорта ---
        # Обработчики уже подключены в _setup_ui при создании QAction
        # --- КОНЕЦ НОВОГО ---

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
        Обработчик импорта. Вызывает соответствующий метод AppController
        на основе выбранных типа и режима.
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

            # Используем выбранные тип и режим
            import_type = self.selected_import_type
            import_mode = self.selected_import_mode
            logger.info(f"Выбран тип импорта: {import_type}, режим: {import_mode} для файла {excel_file_path}")

            # Показываем диалог прогресса
            self.progress_dialog = QProgressDialog("Импорт...", "Отмена", 0, 100, self)
            self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress_dialog.setMinimumDuration(0) # Показываем сразу
            self.progress_dialog.setValue(0)

            # Создаем рабочий поток для вызова AppController метода
            self.import_worker = ImportWorker(self.app_controller, str(excel_file_path), import_type, import_mode)
            self.import_worker.finished.connect(self._on_import_finished)
            # self.import_worker.progress.connect(self._on_import_progress) # Если нужно отображать прогресс изнутри AppController

            self.import_worker.start()

    def _on_import_progress(self, value, message):
        """
        Обработчик прогресса импорта (если поддерживается).
        """
        if self.progress_dialog:
            self.progress_dialog.setValue(value)
            self.progress_dialog.setLabelText(message)
            if self.progress_dialog.wasCanceled():
                # TODO: Реализовать отмену в AppController методах, если возможно
                if self.import_worker:
                    self.import_worker.terminate()
                    self.import_worker.wait()
                self.progress_dialog = None
                self.import_worker = None
                logger.info("Импорт отменён пользователем.")
                self.status_bar.showMessage("Импорт отменён.")
                return

    def _on_import_finished(self, success, message):
        """
        Обработчик завершения импорта.
        """
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None

        if self.import_worker:
            self.import_worker.wait() # Убедиться, что поток завершён
            self.import_worker = None

        if success:
            logger.info(f"Импорт успешно завершён: {message}")
            self.status_bar.showMessage(f"Импорт завершён: {message}")
            # Обновляем список листов, так как могли появиться новые
            self._update_sheet_list()
        else:
            logger.error(f"Импорт завершился с ошибкой: {message}")
            self.status_bar.showMessage(f"Ошибка импорта: {message}")
            QMessageBox.critical(self, "Ошибка", message)

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
        Обновляет список листов в обозревателе.
        """
        # --- НОВОЕ: Обновление через SheetExplorerWidget ---
        self.sheet_explorer.update_sheet_list()
        # --- КОНЕЦ НОВОГО ---

    # --- НОВЫЕ МЕТОДЫ ДЛЯ ОБРАБОТКИ СИГНАЛОВ ОТ SHEET_EXPLORER ---
    def _on_sheet_selected(self, sheet_name: str):
        """
        Обработчик выбора листа в обозревателе.
        """
        if not sheet_name:
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

    def _on_sheet_renamed(self, old_name: str, new_name: str):
        """
        Обработчик успешного переименования листа.
        Может использоваться для обновления UI или логирования.
        Пока просто обновим статусную строку.
        """
        self.status_bar.showMessage(f"Лист переименован: '{old_name}' -> '{new_name}'")
        # Если активный лист был переименован, обновим current_sheet_name
        if self.current_sheet_name == old_name:
            self.current_sheet_name = new_name
    # --- КОНЕЦ НОВЫХ МЕТОДОВ ---

    # --- НОВЫЕ МЕТОДЫ: Обработчики выбора типа и режима, логирования ---
    def _on_logging_toggled(self, state):
        """
        Обработчик переключения состояния чекбокса логирования.
        """
        is_enabled = state == Qt.CheckState.Checked.value
        self.app_controller.set_logging_enabled(is_enabled)
        logger.info(f"Логирование {'включено' if is_enabled else 'отключено'} через GUI.")

    def _on_import_type_selected(self, import_type: str):
        """
        Обработчик выбора типа данных для импорта.
        """
        self.selected_import_type = import_type
        logger.info(f"Выбран тип импорта: {import_type}")
        # Включаем/отключаем режим "Быстрый" в зависимости от типа
        self.import_mode_fast_action.setEnabled(import_type == 'raw_data_pandas')
        # Если выбран 'raw_data_pandas', автоматически выбираем режим 'fast'
        if import_type == 'raw_data_pandas':
            self.import_mode_fast_action.setChecked(True)
            self.selected_import_mode = 'fast'
        # Если был выбран 'fast', но тип изменили не на 'raw_data_pandas', сбрасываем на 'all'
        elif self.selected_import_mode == 'fast' and import_type != 'raw_data_pandas':
            self.import_mode_all_action.setChecked(True)
            self.selected_import_mode = 'all'

    def _on_import_mode_selected(self, import_mode: str):
        """
        Обработчик выбора режима импорта.
        """
        # Проверяем, можно ли выбрать этот режим для текущего типа
        if self.selected_import_type == 'raw_data_pandas' and import_mode != 'fast':
            # Если тип 'raw_data_pandas', но режим не 'fast', сбрасываем на 'fast'
            logger.warning(f"Для типа 'raw_data_pandas' доступен только режим 'fast'. Сброс на 'fast'.")
            self.import_mode_fast_action.setChecked(True)
            self.selected_import_mode = 'fast'
            return
        elif self.selected_import_type != 'raw_data_pandas' and import_mode == 'fast':
            # Если режим 'fast', но тип не 'raw_data_pandas', сбрасываем на 'all'
            logger.warning(f"Режим 'fast' доступен только для типа 'raw_data_pandas'. Сброс на 'all'.")
            self.import_mode_all_action.setChecked(True)
            self.selected_import_mode = 'all'
            return

        self.selected_import_mode = import_mode
        logger.info(f"Выбран режим импорта: {import_mode}")
    # --- КОНЕЦ НОВЫХ МЕТОДОВ ---
