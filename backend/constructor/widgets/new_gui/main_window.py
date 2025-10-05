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
    QCheckBox, QPushButton, QMenu, QDialog # <-- Добавлен QDialog
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QIcon, QActionGroup

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

logger = get_logger(__name__)


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

        # --- НОВОЕ: Подключение чекбокса логирования ---
        self.logging_checkbox.stateChanged.connect(self._on_logging_toggled)
        # --- КОНЕЦ НОВОГО ---

        # --- НОВОЕ: Подключение сигналов обозревателя ---
        self.sheet_explorer.sheet_selected.connect(self._on_sheet_selected)
        self.sheet_explorer.sheet_renamed.connect(self._on_sheet_renamed)
        # --- КОНЕЦ НОВОГО ---

        # --- УДАЛЕНО: Подключение сигналов старого меню типа/режима ---
        # Обработчики уже подключены в _setup_ui при создании QAction
        # --- КОНЕЦ УДАЛЕНИЯ ---

    # --- Обработчики действий меню и тулбара ---
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

    # --- ИЗМЕНЕНО: Новый обработчик импорта ---
    def _on_import_triggered_new(self):
        """
        Обработчик импорта. Открывает ImportDialog.
        """
        logger.info("[НОВЫЙ ДИЗАЙН] Выбрано 'Импорт данных...'")

        # Создаем и показываем диалог импорта
        import_dialog = ImportDialog(self.app_controller, self)
        
        # Показываем модально
        result = import_dialog.exec()
        
        if result == QDialog.DialogCode.Accepted:
            logger.info("[НОВЫЙ ДИЗАЙН] Импорт подтверждён в диалоге.")
            
            # Получаем данные из диалога
            file_path = import_dialog.get_file_path()
            # --- ИЗМЕНЕНО: Получаем ОДИН объединённый ключ режима ---
            import_mode_key = import_dialog.get_import_mode_key()
            # import_type и import_mode больше не существуют как отдельные сущности
            # -------------------------------------------------------
            is_logging_enabled = import_dialog.is_logging_enabled()
            project_name = import_dialog.get_project_name() # Можно использовать позже
            
            if not file_path or not file_path.exists():
                logger.error(f"[НОВЫЙ ДИЗАЙН] Неверный путь к файлу импорта: {file_path}")
                QMessageBox.critical(self, "Ошибка", f"Неверный путь к файлу импорта: {file_path}")
                return

            # Сохраняем выбранный ключ режима для будущего использования
            # (если нужно сохранять между сессиями или для других целей)
            # self.selected_import_mode_key = import_mode_key

            # Устанавливаем состояние логирования через AppController
            self.app_controller.set_logging_enabled(is_logging_enabled)
            logger.info(f"[НОВЫЙ ДИЗАЙН] Логирование {'включено' if is_logging_enabled else 'отключено'} для импорта.")

            # --- ИЗМЕНЕНО: Разбор import_mode_key на import_type и import_mode ---
            parts = import_mode_key.split('_', 1) # Разделить только по первому '_'
            if len(parts) == 2:
                import_type, import_mode = parts
            else:
                # Обработка ошибки или установка значений по умолчанию, если формат неверен
                logger.error(f"[НОВЫЙ ДИЗАЙН] Неверный формат import_mode_key: {import_mode_key}")
                QMessageBox.critical(self, "Ошибка", f"Неверный формат ключа режима импорта: {import_mode_key}")
                return
            # -------------------------------------------------------

            # --- ИЗМЕНЕНО: Передаём import_type и import_mode в ImportWorker ---
            # Создаем рабочий поток для импорта, передав ему ключ режима
            self.import_worker = ImportWorker(self.app_controller, str(file_path), import_type, import_mode)
            # -------------------------------------------------------
            self.import_worker.finished.connect(self._on_import_finished)
            self.import_worker.progress.connect(self._on_import_progress)

            self.import_worker.start()
            self.status_bar.showMessage(f"Начат импорт {file_path.name}...")
            logger.info(f"[НОВЫЙ ДИЗАЙН] Начат импорт: файл={file_path}, тип={import_type}, режим={import_mode}")

        else:
            logger.info("[НОВЫЙ ДИЗАЙН] Импорт отменён пользователем.")

    # --- КОНЕЦ ИЗМЕНЕНИЯ ---

    def _on_import_triggered(self):
        """
        Старый обработчик импорта. Будет заменён на _on_import_triggered_new.
        """
        # Этот метод можно удалить после перехода на новый дизайн
        # или оставить как заглушку/для обратной совместимости
        logger.warning("Вызван устаревший метод _on_import_triggered. Используйте _on_import_triggered_new.")
        self._on_import_triggered_new()

    def _on_import_progress(self, value, message):
        """
        Обработчик прогресса импорта (если поддерживается).
        """
        # Обновляем сообщение в строке состояния
        self.status_bar.showMessage(message)

    def _on_import_finished(self, success, message):
        """
        Обработчик завершения импорта.
        """
        # Убираем ссылку на worker, чтобы он мог быть уничтожен
        if self.import_worker:
            self.import_worker.wait() # Убедиться, что поток завершён
            self.import_worker = None

        if success:
            logger.info(f"[НОВЫЙ ДИЗАЙН] Импорт успешно завершён: {message}")
            self.status_bar.showMessage(f"Импорт завершён: {message}")
            # Обновляем список листов, так как могли появиться новые
            self._update_sheet_list()
        else:
            logger.error(f"[НОВЫЙ ДИЗАЙН] Импорт завершился с ошибкой: {message}")
            self.status_bar.showMessage(f"Ошибка импорта: {message}")
            # QMessageBox.critical(self, "Ошибка", message) # <-- Опционально

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
                # Создаем рабочий поток для вызова AppController метода
                self.export_worker = ExportWorker(self.app_controller, str(output_path))
                self.export_worker.finished.connect(self._on_export_finished)
                self.export_worker.progress.connect(self._on_export_progress) # <-- НОВОЕ: Подключаем progress

                self.export_worker.start()
                # Обновляем статус
                self.status_bar.showMessage(f"Начат экспорт в {output_path.name}...")

            except Exception as e:
                logger.error(f"Ошибка при подготовке экспорта: {e}", exc_info=True)
                QMessageBox.critical(self, "Ошибка", f"Ошибка при подготовке экспорта: {e}")

    def _on_export_progress(self, value, message): # <-- НОВЫЙ МЕТОД
        """
        Обработчик прогресса экспорта.
        """
        # Обновляем сообщение в строке состояния
        self.status_bar.showMessage(message)

    def _on_export_finished(self, success, message):
        """
        Обработчик завершения экспорта.
        """
        # Убираем ссылку на worker, чтобы он мог быть уничтожен
        if self.export_worker:
            self.export_worker.wait() # Убедиться, что поток завершён
            self.export_worker = None

        if success:
            logger.info(f"Экспорт успешно завершён: {message}")
            self.status_bar.showMessage(f"Экспорт завершён: {message}")
        else:
            logger.error(f"Экспорт завершился с ошибкой: {message}")
            self.status_bar.showMessage(f"Ошибка экспорта: {message}")
            # QMessageBox.critical(self, "Ошибка", message) # <-- Опционально, можно оставить

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
        Загружает данные листа в TableEditorWidget.
        """
        if not sheet_name:
            return

        logger.info(f"[НОВЫЙ ДИЗАЙН] Смена активного листа на: {sheet_name}")
        self.current_sheet_name = sheet_name

        # Удаляем предыдущий TableEditorWidget, если он был
        if self.table_editor_widget:
            self.stacked_widget.removeWidget(self.table_editor_widget)
            self.table_editor_widget.deleteLater() # Удаляем из Qt
            self.table_editor_widget = None

        # Создаём новый TableEditorWidget для выбранного листа
        self.table_editor_widget = TableEditorWidget(self.app_controller, self)
        # Загружаем данные листа
        self.table_editor_widget.load_sheet(sheet_name)
        
        self.stacked_widget.addWidget(self.table_editor_widget)
        self.stacked_widget.setCurrentWidget(self.table_editor_widget)

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
        logger.info(f"[НОВЫЙ ДИЗАЙН] Логирование {'включено' if is_enabled else 'отключено'} через GUI.")

    # --- КОНЕЦ НОВЫХ МЕТОДОВ ---