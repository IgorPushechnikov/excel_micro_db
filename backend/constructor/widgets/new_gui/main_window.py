# run_new_gui.py
"""
Точка входа для запуска нового GUI приложения Excel Micro DB.
Следует новому дизайну: таблица - центральный элемент.
"""

import sys
import os
import logging
from pathlib import Path

# --- Добавление корня проекта в sys.path ---
# Это необходимо для корректного импорта модулей из backend
project_root = Path(__file__).parent.resolve()
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))
# -------------------------------------------

# Импортируем QApplication
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QCoreApplication, Qt
from PySide6.QtWidgets import QStyleFactory

# Импортируем новое главное окно
# Убедимся, что импортируем из правильного модуля
# Если main_window_new.py был переименован в main_window.py, путь будет другой
# from backend.constructor.widgets.new_gui.main_window_new import MainWindowNew
# Предположим, что мы будем использовать обновлённый main_window.py
# from backend.constructor.widgets.new_gui.main_window import MainWindow

# Импортируем логгер
from backend.utils.logger import get_logger, setup_logger

logger = get_logger(__name__)

def main():
    """
    Основная функция запуска нового GUI.
    """
    # Настройка логирования
    # setup_logger() вызывается внутри MainWindow, но можно и здесь для раннего лога
    setup_logger() 
    logger.info("Запуск НОВОГО GUI приложения Excel Micro DB...")

    try:
        # Создание экземпляра QApplication
        app = QApplication(sys.argv)
        
        # Установка имени и версии приложения
        QCoreApplication.setApplicationName("Excel Micro DB New GUI")
        QCoreApplication.setOrganizationName("ExcelMicroDB")
        QCoreApplication.setApplicationVersion("0.2.0")
        
        # Установка стиля (опционально, для лучшего внешнего вида)
        app.setStyle(QStyleFactory.create("Fusion"))

        # Создание и отображение главного окна
        # Используем обновлённый MainWindow, который теперь следует новому дизайну
        window = MainWindow()
        window.show()

        logger.info("НОВОЕ GUI приложение запущено.")

        # Запуск цикла событий Qt
        exit_code = app.exec()
        logger.info(f"НОВОЕ GUI приложение завершено с кодом: {exit_code}")

        sys.exit(exit_code)

    except Exception as e:
        logger.critical(f"Критическая ошибка при запуске НОВОГО GUI: {e}", exc_info=True)
        # Используем стандартный print, так как QApplication может не быть создан
        print(f"[КРИТИЧЕСКАЯ ОШИБКА] Произошла критическая ошибка при запуске приложения:\n{e}\n\nПриложение будет закрыто.", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()



# backend/constructor/widgets/new_gui/main_window.py
"""
Главное окно приложения Excel Micro DB GUI.
Управляет проектами, листами и отображает TableEditorWidget (новый дизайн).
"""

import logging
from pathlib import Path
from typing import Optional, Callable # <-- Добавлен Callable
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QStatusBar, QToolBar, QFileDialog,
    QMessageBox, QWidget, QStackedWidget, QHBoxLayout, QSplitter,
    QCheckBox, QPushButton, QMenu, QDialog, QProgressDialog, QProgressBar # <-- Добавлен QProgressBar
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer # <-- Добавлен QTimer
from PySide6.QtGui import QAction, QIcon, QActionGroup, QCursor # <-- Добавлен QCursor

from backend.core.app_controller import create_app_controller
# Убираем старый импорт
# from .sheet_editor_widget import SheetEditorWidget
# Добавляем новый импорт TableEditorWidget
from .table_editor_widget_new import TableEditorWidget
from .import_worker import ImportWorker # <-- НОВЫЙ ИМПОРТ
# Добавляем импорт ImportDialog
from .import_dialog_new import ImportDialog
from .sheet_explorer_widget import SheetExplorerWidget
from .export_worker import ExportWorker
from backend.utils.logger import get_logger
# --- НОВЫЙ ИМПОРТ ---
from backend.importer.xlwings_importer import import_all_from_excel_xlwings

logger = get_logger(__name__)


# --- НОВОЕ: Класс потока для xlwings-импорта ---
class XlImportThread(QThread):
    """
    Поток для выполнения импорта через xlwings.
    """
    # Сигнал прогресса: (процент, сообщение)
    progress = Signal(int, str)
    # Сигнал завершения: (успех, сообщение)
    finished = Signal(bool, str)

    def __init__(self, db_path, file_path):
        """
        Args:
            db_path (str): Путь к файлу БД проекта.
            file_path (str): Путь к Excel-файлу
        """
        super().__init__()
        self.db_path = db_path
        self.file_path = file_path

    def run(self):
        """
        Запуск импорта в отдельном потоке.
        """
        from backend.storage.base import ProjectDBStorage
        storage = ProjectDBStorage(self.db_path)
        if not storage.connect():
            logger.error(f"XlImportThread: Не удалось подключиться к БД проекта {self.db_path}.")
            self.finished.emit(False, f"Не удалось подключиться к БД: {self.db_path}")
            return

        try:
            # Передаём функцию обновления прогресса
            success = import_all_from_excel_xlwings(
                storage,
                self.file_path,
                progress_callback=self._on_progress
            )
            if success:
                self.finished.emit(True, "Импорт завершён успешно.")
            else:
                self.finished.emit(False, "Ошибка при импорте через xlwings.")
        except Exception as e:
            logger.error(f"Ошибка в потоке xlwings-импорта: {e}", exc_info=True)
            self.finished.emit(False, f"Ошибка: {e}")
        finally:
            storage.disconnect()

    def _on_progress(self, percent: int, message: str):
        """
        Callback для обновления прогресса.
        """
        self.progress.emit(percent, message)
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
        # --- ИЗМЕНЕНО: Атрибут для хранения TableEditorWidget ---
        # self.sheet_editor_widget = None
        self.table_editor_widget: Optional[TableEditorWidget] = None
        # --- КОНЕЦ ИЗМЕНЕНИЯ ---
        self.import_worker = None
        self.export_worker = None
        # --- НОВОЕ: Атрибут для потока xlwings-импорта ---
        self.xl_import_thread: Optional[XlImportThread] = None
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
        # --- ИЗМЕНЕНО: Импорт теперь через диалог ---
        # self.import_action = QAction("Импорт", self) # Основное действие
        self.import_action = QAction("Импорт данных...", self) # Более конкретное название
        # --- КОНЕЦ ИЗМЕНЕНИЯ ---
        self.export_action = QAction("Экспорт", self)
        # --- НОВОЕ: Действие для xlwings ---
        self.import_xlwings_action = QAction("Забрать из Excel", self)
        # --- КОНЕЦ НОВОГО ---

        file_menu.addAction(self.new_project_action)
        file_menu.addAction(self.open_project_action)
        file_menu.addAction(self.save_project_action)
        import_menu.addAction(self.import_action)
        # --- ДОБАВЛЕНО: Добавляем xlwings-импорт в меню ---
        import_menu.addAction(self.import_xlwings_action)
        # --- КОНЕЦ ДОБАВЛЕНИЯ ---
        export_menu.addAction(self.export_action)

        # Панель инструментов
        tool_bar = QToolBar("Основная", self)
        self.addToolBar(tool_bar)

        # --- НОВОЕ: Чекбокс для управления логированием ---
        self.logging_checkbox = QCheckBox("Логирование", self)
        self.logging_checkbox.setChecked(False)  # По умолчанию отключено
        tool_bar.addWidget(self.logging_checkbox)
        # --- КОНЕЦ НОВОГО ---

        # --- УДАЛЕНО: Старая кнопка и меню для типа/режима импорта ---
        # self.import_type_button = QPushButton("Тип импорта", self)
        # self.import_type_menu = QMenu(self.import_type_button)
        # ... (весь старый код меню)
        # tool_bar.addWidget(self.import_type_button)
        # --- КОНЕЦ УДАЛЕНИЯ ---

        # Статусная строка
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов.")
        # --- НОВОЕ: Добавление QProgressBar в status_bar ---
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False) # Скрыта по умолчанию
        self.status_bar.addPermanentWidget(self.progress_bar)
        # --- КОНЕЦ НОВОГО ---

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
        # --- ИЗМЕНЕНО: Подключение нового обработчика импорта ---
        # self.import_action.triggered.connect(self._on_import_triggered)
        self.import_action.triggered.connect(self._on_import_triggered_new)
        # --- КОНЕЦ ИЗМЕНЕНИЯ ---
        self.export_action.triggered.connect(self._on_export_triggered)
        # --- НОВОЕ: Подключение xlwings-импорта ---
        self.import_xlwings_action.triggered.connect(self._on_import_xlwings_triggered)
        # --- КОНЕЦ НОВОГО ---

        # --- НОВОЕ: Подключение чекбокса логирования ---
        self.logging_checkbox.stateChanged.connect(self._on_logging_toggled)
        # --- КОНЕЦ НОВОГО ---

        # --- НОВОЕ: Подключение сигналов обозревателя ---
        self.sheet_explorer.sheet_selected.connect(self._on_sheet_selected)
        self.sheet_explorer.sheet_renamed.connect(self._on_sheet_renamed)
        # --- КОНЕЦ НОВОГО ---

    # --- НОВОЕ: Методы для управления прогресс-баром ---
    def set_progress(self, value: int, message: str = ""):
        """
        Обновляет значение и текст прогресс-бара в статусной строке.

        Args:
            value (int): Значение прогресса (0-100).
            message (str): Текстовое сообщение для отображения.
        """
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(f"{message} %p%") # %p% - встроенный процент
        if not self.progress_bar.isVisible():
            self.progress_bar.setVisible(True)

    def hide_progress(self):
        """
        Скрывает прогресс-бар в статусной строке.
        """
        self.progress_bar.setVisible(False)
        self.status_bar.showMessage("Готов.") # Возвращаем сообщение "Готов" при скрытии
    # --- КОНЕЦ НОВОГО ---

    # --- НОВОЕ: Обработчик xlwings-импорта ---
    def _on_import_xlwings_triggered(self):
        """
        Обработчик для действия "Забрать из Excel".
        """
        if not self.current_project_path:
            QMessageBox.warning(self, "Нет проекта", "Пожалуйста, сначала создайте или откройте проект.")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel-файл для импорта",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if not file_path:
            return

        if not self.app_controller or not self.app_controller.project_db_path:
            logger.error("AppController или его project_db_path не инициализированы.")
            QMessageBox.critical(self, "Ошибка", "AppController не инициализирован или проект не загружен.")
            return

        # --- НОВОЕ: Создаём поток ---
        # Передаём db_path вместо storage
        self.xl_import_thread = XlImportThread(self.app_controller.project_db_path, file_path)
        # Подключаем сигнал progress к методу обновления прогресса в MainWindow
        self.xl_import_thread.progress.connect(self.set_progress)
        self.xl_import_thread.finished.connect(lambda success, msg: (
            self.hide_progress(), # Скрываем прогресс после завершения
            QMessageBox.information(self, "Успех", msg) if success else QMessageBox.critical(self, "Ошибка", msg),
            self.sheet_explorer.update_sheet_list() if success else None
        ))
        self.xl_import_thread.start()
        # --- КОНЕЦ НОВОГО ---

    # --- КОНЕЦ НОВОГО ---

    def _on_new_project_triggered(self):
        """
        Обработчик создания нового проекта.
        """
        # --- ИСПРАВЛЕНО: Используем getExistingDirectory для выбора директории ---
        project_dir_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите директорию для нового проекта",
            ""
        )
        if not project_dir_path:
            return

        # Используем basename директории как имя проекта
        project_name = Path(project_dir_path).name

        # --- ИСПРАВЛЕНО: вызов create_project с ОДНИМ аргументом (директорией) ---
        if self.app_controller.create_project(project_dir_path):
            self.current_project_path = project_dir_path
            self.status_bar.showMessage(f"Проект создан: {project_dir_path}")
            # --- НОВОЕ: Явно загружаем проект, чтобы инициализировать storage ---
            if not self.app_controller.load_project(project_dir_path):
                logger.error("Не удалось загрузить только что созданный проект.")
                # Можно показать сообщение пользователю
                QMessageBox.critical(self, "Ошибка", "Проект создан, но не удалось его загрузить.")
                return
            # --- КОНЕЦ НОВОГО ---
            self.sheet_explorer.update_sheet_list()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось создать проект.")

    def _on_open_project_triggered(self):
        """
        Обработчик открытия проекта.
        """
        # --- ИСПРАВЛЕНО: Используем getExistingDirectory для выбора директории ---
        project_dir_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите директорию проекта",
            ""
        )
        if not project_dir_path:
            return

        # --- ИСПРАВЛЕНО: вызов load_project с ОДНИМ аргументом (директорией) ---
        if self.app_controller.load_project(project_dir_path):
            self.current_project_path = project_dir_path
            self.status_bar.showMessage(f"Проект открыт: {project_dir_path}")
            self.sheet_explorer.update_sheet_list()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось открыть проект.")

    def _on_save_project_triggered(self):
        """
        Обработчик сохранения проекта.
        """
        # --- ИСПРАВЛЕНО: save_project не существует, просто обновляем статус ---
        if not self.current_project_path:
            QMessageBox.warning(self, "Нет проекта", "Пожалуйста, сначала создайте или откройте проект.")
            return
        # AppController не имеет метода save_project. Данные сохраняются при закрытии/работе с БД.
        self.status_bar.showMessage(f"Проект сохранён: {self.current_project_path}")
        # self.sheet_explorer.update_sheet_list() # <-- Опционально, если данные обновлены

    def _on_import_triggered_new(self):
        """
        Обработчик импорта данных через диалог.
        """
        if not self.current_project_path:
            QMessageBox.warning(self, "Нет проекта", "Пожалуйста, сначала создайте или откройте проект.")
            return

        dialog = ImportDialog(self.app_controller, self) # <-- ИСПРАВЛЕНО: передаём app_controller
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # --- ИСПРАВЛЕНО: вызов методов диалога ---
            file_path = dialog.get_file_path()
            import_mode_key = dialog.get_import_mode_key()
            # --- НОВОЕ: Получаем состояние чекбокса "Выборочно" ---
            is_selective_import = dialog.is_selective_import_checked()
            # --- КОНЕЦ НОВОГО ---
            # Разбор ключа на тип и режим (если нужно)
            # import_type, import_mode = import_mode_key.split('_') # <-- УДАЛЕНО: теперь ключ объединён
            # --- КОНЕЦ ИСПРАВЛЕНИЯ ---

            # --- НОВОЕ: Обработка выборочного импорта ---
            selective_options = None
            if is_selective_import:
                logger.info("Запущен режим выборочного импорта. Получение списка листов...")
                
                # 1. Определяем тип импорта (openpyxl или xlwings) из import_mode_key
                # Пример: 'all_openpyxl', 'raw_openpyxl', 'all_xlwings'
                if 'xlwings' in import_mode_key:
                    import_method = 'xlwings'
                elif 'openpyxl' in import_mode_key:
                    import_method = 'openpyxl'
                else:
                    # По умолчанию, если не указано, предположим openpyxl
                    import_method = 'openpyxl'
                    logger.warning(f"Не удалось определить метод импорта из ключа '{import_mode_key}'. Используется 'openpyxl' по умолчанию.")

                # 2. Получаем список листов
                available_sheet_names = []
                try:
                    if import_method == 'openpyxl':
                        import openpyxl
                        logger.debug(f"Открытие файла '{file_path}' через openpyxl (read_only=True) для получения списка листов...")
                        # --- НОВОЕ: Обработка ошибки Nested.from_tree ---
                        try:
                            wb_temp = openpyxl.load_workbook(str(file_path), read_only=True, data_only=True)
                        except TypeError as e:
                            if "Nested.from_tree() missing 1 required positional argument: 'node'" in str(e):
                                logger.error(f"Ошибка openpyxl при открытии файла '{file_path}' для получения списка листов: {e}")
                                QMessageBox.critical(self, "Ошибка", f"Файл '{file_path}' содержит неподдерживаемые структуры (например, pivot-таблицы) и не может быть обработан. Ошибка: {e}")
                                return # Прерываем импорт
                            else:
                                raise # Если это другая ошибка TypeError, пробрасываем её
                        # --- КОНЕЦ НОВОГО ---
                        available_sheet_names = wb_temp.sheetnames
                        wb_temp.close()
                        logger.debug(f"Получен список листов (openpyxl): {available_sheet_names}")
                    elif import_method == 'xlwings':
                        import xlwings as xw
                        logger.debug(f"Открытие файла '{file_path}' через xlwings (visible=False, read_only=True) для получения списка листов...")
                        app_temp = xw.App(visible=False)
                        wb_temp = app_temp.books.open(str(file_path), update_links=False, read_only=True)
                        available_sheet_names = [s.name for s in wb_temp.sheets]
                        wb_temp.close()
                        app_temp.quit()
                        logger.debug(f"Получен список листов (xlwings): {available_sheet_names}")
                except Exception as e:
                    logger.error(f"Ошибка при получении списка листов: {e}", exc_info=True)
                    QMessageBox.critical(self, "Ошибка", f"Не удалось получить список листов из файла: {e}")
                    return # Прерываем импорт
                
                if not available_sheet_names:
                    logger.warning("Файл не содержит листов или список листов пуст.")
                    QMessageBox.warning(self, "Предупреждение", "Файл не содержит листов или не удалось их получить.")
                    return # Прерываем импорт

                # 3. Показываем диалог выбора листов
                logger.debug("Показ диалога выбора листов для выборочного импорта...")
                from .selective_import_options_dialog import SelectiveImportOptionsDialog
                select_dialog = SelectiveImportOptionsDialog(available_sheet_names, self)
                if select_dialog.exec() == QDialog.DialogCode.Accepted:
                    selected_sheet_names = select_dialog.get_selected_sheet_names()
                    logger.info(f"Выбраны листы для импорта: {selected_sheet_names}")
                    selective_options = {'sheets': selected_sheet_names}
                    # Можно добавить другие опции позже, например, 'start_row', 'end_row'
                else:
                    logger.info("Диалог выбора листов отменён пользователем. Импорт прерван.")
                    return # Прерываем импорт, если диалог отменён
            
            # --- КОНЕЦ НОВОГО ---

            # --- ИСПРАВЛЕНО: вызов ImportWorker с правильными аргументами ---
            # self.import_worker = ImportWorker(self.app_controller, str(file_path), import_mode_key) # <-- СТАРОЕ
            self.import_worker = ImportWorker(self.app_controller, str(file_path), import_mode_key, selective_options) # <-- НОВОЕ
            # Подключаем сигнал progress к методу обновления прогресса в MainWindow
            self.import_worker.progress.connect(self.set_progress)
            # Подключаем сигнал finished вместо import_finished
            self.import_worker.finished.connect(lambda success, msg: (
                self.hide_progress(), # Скрываем прогресс после завершения
                QMessageBox.information(self, "Успех", msg) if success else QMessageBox.critical(self, "Ошибка", msg),
                self.sheet_explorer.update_sheet_list() if success else None # Обновляем список листов при успехе
            ))
            self.import_worker.start()
        else:
            logger.debug("Диалог импорта отменён пользователем.")

    def _on_import_finished(self, success: bool, message: str): # <-- ИСПРАВЛЕНО: подпись сигнала
        """
        Обработчик завершения импорта.
        """
        # Этот метод больше не нужен, так как подключение к finished происходит в _on_import_triggered_new
        pass

    def _on_export_triggered(self):
        """
        Обработчик экспорта данных.
        """
        if not self.current_project_path:
            QMessageBox.warning(self, "Нет проекта", "Пожалуйста, сначала создайте или откройте проект.")
            return

        # --- ИСПРАВЛЕНО: вызов ExportWorker с правильными аргументами ---
        # Пусть ExportWorker сам решает, куда экспортировать, или запросит через QFileDialog
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Экспорт проекта",
            "",
            "Excel Files (*.xlsx)"
        )
        if not output_path:
            return

        # Запускаем экспорт в отдельном потоке
        self.export_worker = ExportWorker(self.app_controller, output_path)
        # Подключаем сигнал progress к методу обновления прогресса в MainWindow
        self.export_worker.progress.connect(self.set_progress)
        # Подключаем сигнал finished вместо export_finished
        self.export_worker.finished.connect(lambda success, msg: (
            self.hide_progress(), # Скрываем прогресс после завершения
            QMessageBox.information(self, "Успех", msg) if success else QMessageBox.critical(self, "Ошибка", msg)
        ))
        self.export_worker.start()

    def _on_export_finished(self, success: bool, message: str): # <-- ИСПРАВЛЕНО: подпись сигнала
        """
        Обработчик завершения экспорта.
        """
        # Этот метод больше не нужен, так как подключение к finished происходит в _on_export_triggered
        pass

    def _on_logging_toggled(self, state):
        """
        Обработчик переключения логирования.
        """
        # state: 0 = Qt.Unchecked, 2 = Qt.Checked
        enabled = state == Qt.CheckState.Checked.value
        # Включаем/выключаем логирование через AppController
        self.app_controller.set_logging_enabled(enabled)

        if enabled:
            self.status_bar.showMessage("Логирование включено.")
            # Логируем через локальный логгер, он тоже будет подчиняться общему правилу
            logger.debug("Логирование включено через GUI.")
        else:
            self.status_bar.showMessage("Логирование отключено.")
            # Это сообщение не появится в консоли, если консольный уровень excel_micro_db > INFO
            logger.debug("Логирование отключено через GUI.")

    def _on_sheet_selected(self, sheet_name: str):
        """
        Обработчик выбора листа в обозревателе.
        """
        self.current_sheet_name = sheet_name
        logger.debug(f"Выбран лист: {sheet_name}")

        # Проверяем, существует ли уже редактор для этого листа
        for i in range(self.stacked_widget.count()):
            widget = self.stacked_widget.widget(i)
            # --- ИСПРАВЛЕНО: проверка атрибута sheet_name ---
            # hasattr(widget, 'sheet_name') and widget.sheet_name == sheet_name:
            # Нет атрибута sheet_name, используем метод из модели
            if isinstance(widget, TableEditorWidget):
                current_sheet = widget.get_current_sheet_name()
                if current_sheet and current_sheet == sheet_name:
                    self.stacked_widget.setCurrentIndex(i)
                    return
            # --- КОНЕЦ ИСПРАВЛЕНИЯ ---

        # Создаём новый редактор
        table_editor = TableEditorWidget(self.app_controller, self)
        # table_editor.sheet_name = sheet_name  # Простое поле для идентификации - УДАЛЕНО
        table_editor.load_sheet(sheet_name) # <-- Загружаем данные листа
        index = self.stacked_widget.addWidget(table_editor)
        self.stacked_widget.setCurrentIndex(index)

    def _on_sheet_renamed(self, old_name: str, new_name: str):
        """
        Обработчик переименования листа.
        """
        logger.info(f"Лист переименован: {old_name} -> {new_name}")
        # Обновляем имя в TableEditorWidget, если он открыт
        for i in range(self.stacked_widget.count()):
            widget = self.stacked_widget.widget(i)
            if isinstance(widget, TableEditorWidget):
                current_sheet = widget.get_current_sheet_name()
                if current_sheet and current_sheet == old_name:
                    # widget.sheet_name = new_name # <-- УДАЛЕНО: нет атрибута
                    # widget.load_data() # Пример: перезагрузить данные под новым именем - УДАЛЕНО: нет метода
                    widget.load_sheet(new_name) # <-- Правильный способ перезагрузить данные

# --- Конец класса MainWindow ---
