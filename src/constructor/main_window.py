# src/constructor/main_window.py
"""
Главное окно графического интерфейса Excel Micro DB.
"""
import sys
from pathlib import Path
from typing import Optional

# Импорт Qt
from PySide6.QtWidgets import (
    QMainWindow, QMenu, QMenuBar, QStatusBar, QWidget, QVBoxLayout, QLabel,
    QFileDialog, QMessageBox, QApplication, QDockWidget, QStackedWidget,
    QProgressDialog # НОВОЕ: для отображения прогресса
)
from PySide6.QtCore import Qt, Slot, QTimer # QTimer для имитации асинхронности
from PySide6.QtGui import QAction, QCloseEvent
import logging

# Импорт из проекта
from src.core.app_controller import create_app_controller, AppController
from src.utils.logger import get_logger
# Импорт новых виджетов
from src.constructor.widgets.project_explorer import ProjectExplorer
from src.constructor.widgets.sheet_editor import SheetEditor

logger = get_logger(__name__)

class MainWindow(QMainWindow):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()
        self.app_controller: Optional[AppController] = None
        self.project_explorer: Optional[ProjectExplorer] = None
        self.sheet_editor: Optional[SheetEditor] = None
        self.central_stacked_widget: Optional[QStackedWidget] = None
        self._welcome_widget: Optional[QWidget] = None
        # НОВОЕ: Диалог прогресса
        self.progress_dialog: Optional[QProgressDialog] = None
        self._setup_ui()
        self._setup_controller()
        self._update_ui_state()
        logger.info("MainWindow инициализировано")

    def _setup_ui(self):
        """Настройка пользовательского интерфейса."""
        self.setWindowTitle("Excel Micro DB - Конструктор")
        self.resize(1000, 700)

        # --- Меню ---
        menubar = self.menuBar()
        self.file_menu = menubar.addMenu('&Файл')
        
        self.action_new_project = QAction('&Создать проект...', self)
        self.action_new_project.setShortcut('Ctrl+N')
        self.action_new_project.triggered.connect(self._on_new_project)
        self.file_menu.addAction(self.action_new_project)
        
        self.action_open_project = QAction('&Открыть проект...', self)
        self.action_open_project.setShortcut('Ctrl+O')
        self.action_open_project.triggered.connect(self._on_open_project)
        self.file_menu.addAction(self.action_open_project)
        
        self.file_menu.addSeparator()
        
        # НОВОЕ: Действие "Анализировать файл"
        self.action_analyze_file = QAction('&Анализировать файл...', self)
        self.action_analyze_file.setShortcut('Ctrl+A')
        self.action_analyze_file.triggered.connect(self._on_analyze_file)
        # Изначально отключено, будет включено при загрузке проекта
        self.action_analyze_file.setEnabled(False)
        self.file_menu.addAction(self.action_analyze_file)
        
        self.file_menu.addSeparator()
        
        self.action_exit = QAction('&Выход', self)
        self.action_exit.setShortcut('Ctrl+Q')
        self.action_exit.triggered.connect(self.close)
        self.file_menu.addAction(self.action_exit)

        # --- Центральный виджет (Стек) ---
        self.central_stacked_widget = QStackedWidget()
        self.setCentralWidget(self.central_stacked_widget)

        self._welcome_widget = QWidget()
        welcome_layout = QVBoxLayout(self._welcome_widget)
        self.welcome_label = QLabel("Добро пожаловать в Excel Micro DB!\n\n"
                                   "1. Создайте новый проект (Файл -> Создать проект)\n"
                                   "2. Или откройте существующий (Файл -> Открыть проект)\n"
                                   "3. После загрузки проекта можно анализировать Excel-файлы (Файл -> Анализировать файл)")
        self.welcome_label.setAlignment(Qt.AlignCenter)
        self.welcome_label.setWordWrap(True)
        welcome_layout.addWidget(self.welcome_label)
        self.central_stacked_widget.addWidget(self._welcome_widget)

        self.sheet_editor = SheetEditor()
        # === НОВОЕ: Передача AppController в SheetEditor ===
        # Это необходимо для того, чтобы SheetEditor мог вызывать методы
        # AppController для загрузки редактируемых данных и сохранения изменений.
        # Пока контроллер может быть None, но он будет установлен позже в _setup_controller
        # и обновлён в _update_ui_state если проект загружается/создаётся.
        self.sheet_editor.set_app_controller(self.app_controller)
        # =====================================================
        self.central_stacked_widget.addWidget(self.sheet_editor)
        self.central_stacked_widget.setCurrentWidget(self._welcome_widget)

        # --- Обозреватель проекта ---
        self.project_explorer = ProjectExplorer(self)
        self.project_explorer.sheet_selected.connect(self._on_sheet_selected)
        self.addDockWidget(Qt.LeftDockWidgetArea, self.project_explorer)
        
        # --- Статусная строка ---
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готово", 5000)

    def _setup_controller(self):
        """Инициализация контроллера приложения."""
        try:
            logger.debug("Создание AppController")
            self.app_controller = create_app_controller()
            init_success = self.app_controller.initialize()
            if not init_success:
                raise Exception("Не удалось инициализировать AppController")
            logger.info("AppController инициализирован и готов к работе")
            # === НОВОЕ: Обновление контроллера в SheetEditor после инициализации ===
            if self.sheet_editor:
                 self.sheet_editor.set_app_controller(self.app_controller)
            # ======================================================================
        except Exception as e:
            logger.error(f"Ошибка при инициализации AppController: {e}")
            QMessageBox.critical(
                self,
                "Ошибка инициализации",
                f"Не удалось инициализировать ядро приложения:\n{e}\n\n"
                "Приложение может работать некорректно."
            )
            self.app_controller = None

    def _update_ui_state(self):
        """Обновление состояния UI в зависимости от состояния контроллера."""
        controller_ready = self.app_controller is not None
        project_loaded = controller_ready and self.app_controller.is_project_loaded
        
        self.action_new_project.setEnabled(controller_ready)
        self.action_open_project.setEnabled(controller_ready)
        # НОВОЕ: Включаем/выключаем действие анализа
        self.action_analyze_file.setEnabled(project_loaded)
        
        if not controller_ready:
             self.status_bar.showMessage("Ошибка: Контроллер приложения не доступен", 0)
             
        if self.project_explorer:
            if project_loaded:
                project_data = self.app_controller.current_project
                project_path = self.app_controller.project_path
                if project_data and project_path:
                    db_path = project_path / "project_data.db"
                    self.project_explorer.load_project(project_data, str(db_path))
                else:
                    self.project_explorer.clear_project()
            else:
                self.project_explorer.clear_project()
        
        # === НОВОЕ: Обновление контроллера в SheetEditor при изменении состояния ===
        # Это гарантирует, что SheetEditor всегда имеет актуальную ссылку на AppController,
        # особенно важно после загрузки/создания проекта или в случае ошибки контроллера.
        if self.sheet_editor:
            self.sheet_editor.set_app_controller(self.app_controller)
        # ===========================================================================

    @Slot()
    def _on_new_project(self):
        """Обработчик действия 'Создать проект'."""
        logger.info("Начало создания нового проекта")
        if not self.app_controller:
            QMessageBox.warning(self, "Ошибка", "Контроллер приложения не инициализирован.")
            return

        project_dir = QFileDialog.getExistingDirectory(
            self, "Выберите директорию для нового проекта"
        )
        if not project_dir:
            logger.info("Создание проекта отменено пользователем")
            return

        try:
            project_name = Path(project_dir).name
            success = self.app_controller.create_project(project_dir, project_name)
            
            if success:
                self.status_bar.showMessage(f"Проект '{project_name}' создан и загружен", 5000)
                QMessageBox.information(self, "Успех", f"Проект '{project_name}' успешно создан в {project_dir}")
                self._update_ui_state()
                self.central_stacked_widget.setCurrentWidget(self._welcome_widget)
            else:
                self.status_bar.showMessage("Ошибка при создании проекта", 0)
                QMessageBox.critical(self, "Ошибка", "Не удалось создать проект. Подробности в логе.")
        except Exception as e:
            logger.error(f"Необработанная ошибка при создании проекта: {e}", exc_info=True)
            self.status_bar.showMessage("Необработанная ошибка при создании проекта", 0)
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {e}")

    @Slot()
    def _on_open_project(self):
        """Обработчик действия 'Открыть проект'."""
        logger.info("Начало открытия проекта")
        if not self.app_controller:
            QMessageBox.warning(self, "Ошибка", "Контроллер приложения не инициализирован.")
            return

        project_dir = QFileDialog.getExistingDirectory(
            self, "Выберите директорию проекта"
        )
        if not project_dir:
            logger.info("Открытие проекта отменено пользователем")
            return

        try:
            success = self.app_controller.load_project(project_dir)
            
            if success:
                project_name = self.app_controller.current_project.get("project_name", "Неизвестный проект")
                self.status_bar.showMessage(f"Проект '{project_name}' загружен", 5000)
                QMessageBox.information(self, "Успех", f"Проект '{project_name}' успешно загружен из {project_dir}")
                self._update_ui_state()
                self.central_stacked_widget.setCurrentWidget(self._welcome_widget)
            else:
                self.status_bar.showMessage("Ошибка при загрузке проекта", 0)
                QMessageBox.critical(self, "Ошибка", "Не удалось загрузить проект. Убедитесь, что это корректная директория проекта.")
        except Exception as e:
            logger.error(f"Необработанная ошибка при открытии проекта: {e}", exc_info=True)
            self.status_bar.showMessage("Необработанная ошибка при открытии проекта", 0)
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {e}")

    @Slot()
    def _on_analyze_file(self):
        """Обработчик действия 'Анализировать файл'."""
        logger.info("Начало анализа Excel-файла")
        if not self.app_controller or not self.app_controller.is_project_loaded:
            QMessageBox.warning(self, "Ошибка", "Проект не загружен.")
            return

        # Диалог выбора Excel-файла
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel-файл для анализа", "", "Excel Files (*.xlsx *.xls)"
        )
        if not file_path:
            logger.info("Анализ файла отменён пользователем")
            return

        try:
            # Создаем и настраиваем диалог прогресса
            # Примечание: анализ может быть быстрым, диалог может моргнуть.
            # Для реального прогресса нужно модифицировать analyze_excel_file
            # чтобы он мог сообщать о прогрессе.
            self.progress_dialog = QProgressDialog("Анализ файла...", "Отмена", 0, 0, self)
            self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress_dialog.setWindowTitle("Анализ")
            # Скрываем кнопку отмены, так как analyze_excel_file не поддерживает отмену
            self.progress_dialog.setCancelButton(None)
            self.progress_dialog.show()

            # Принудительное обновление UI до начала длительной операции
            QApplication.processEvents()

            # Вызов анализа через контроллер
            # analyze_excel_file может быть блокирующим, поэтому диалог будет "зависшим"
            # В реальном приложении лучше запускать это в отдельном потоке.
            success = self.app_controller.analyze_excel_file(file_path)

            # Закрываем диалог прогресса
            if self.progress_dialog:
                self.progress_dialog.close()
                self.progress_dialog = None

            if success:
                self.status_bar.showMessage("Анализ файла завершён успешно", 5000)
                QMessageBox.information(self, "Успех", f"Файл '{Path(file_path).name}' успешно проанализирован.")
                # Обновляем обозреватель проекта, чтобы отобразить новые данные
                self._update_ui_state()
                # Переключаемся на приветствие/пустой редактор
                self.central_stacked_widget.setCurrentWidget(self._welcome_widget)
            else:
                self.status_bar.showMessage("Ошибка при анализе файла", 0)
                QMessageBox.critical(self, "Ошибка", "Не удалось проанализировать файл. Подробности в логе.")

        except Exception as e:
            # Закрываем диалог прогресса в случае ошибки
            if self.progress_dialog:
                self.progress_dialog.close()
                self.progress_dialog = None

            logger.error(f"Необработанная ошибка при анализе файла: {e}", exc_info=True)
            self.status_bar.showMessage("Необработанная ошибка при анализе файла", 0)
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при анализе файла:\n{e}")

    @Slot(str)
    def _on_sheet_selected(self, sheet_name: str):
        """
        Обработчик сигнала sheet_selected от ProjectExplorer.
        Вызывается, когда пользователь выбирает лист в обозревателе.
        """
        logger.debug(f"MainWindow получил сигнал о выборе листа: {sheet_name}")
        # === ИЗМЕНЕНО: Логика загрузки теперь внутри SheetEditor.load_sheet ===
        # SheetEditor теперь сам взаимодействует с AppController
        if self.app_controller and self.app_controller.is_project_loaded and self.app_controller.project_path:
            try:
                db_path = self.app_controller.project_path / "project_data.db"
                self.central_stacked_widget.setCurrentWidget(self.sheet_editor)
                # SheetEditor.load_sheet теперь использует AppController для получения данных
                self.sheet_editor.load_sheet(str(db_path), sheet_name)
                self.status_bar.showMessage(f"Открыт лист: {sheet_name}", 3000)
            except Exception as e:
                logger.error(f"Ошибка при открытии листа '{sheet_name}' в редакторе: {e}", exc_info=True)
                self.status_bar.showMessage(f"Ошибка при открытии листа: {e}", 0)
                QMessageBox.critical(self, "Ошибка", f"Не удалось открыть лист '{sheet_name}':\n{e}")
                self.central_stacked_widget.setCurrentWidget(self._welcome_widget)
        else:
            logger.warning("Попытка открыть лист, но проект не загружен.")
            self.status_bar.showMessage("Ошибка: Проект не загружен.", 0)
            QMessageBox.warning(self, "Ошибка", "Проект не загружен. Невозможно открыть лист.")
        # ======================================================================

    def closeEvent(self, event: QCloseEvent):
        """Обработчик события закрытия окна."""
        logger.info("Получен запрос на закрытие главного окна")
        # Закрываем диалог прогресса, если он открыт
        if self.progress_dialog:
            self.progress_dialog.close()
        try:
            if self.app_controller:
                self.app_controller.shutdown()
        except Exception as e:
            logger.error(f"Ошибка при завершении работы контроллера: {e}", exc_info=True)
        finally:
            event.accept()
            logger.info("Главное окно закрыто")
