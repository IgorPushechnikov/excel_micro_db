# src/constructor_new/main_window.py
"""
Модуль для главного окна нового GUI Excel Micro DB.
Использует PySide6, приближен к интерфейсу Excel.
"""

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QTabWidget,
    QMenuBar, QToolBar, QStatusBar, QDockWidget, QTableView, QHeaderView,
    QAbstractItemView, QFrame, QLabel, QLineEdit, QToolButton
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QAction, QIcon

# Импорты для нового GUI
# from .components.sheet_editor import SheetEditor  # Пока нет
# from .components.project_explorer import ProjectExplorer  # Пока нет
# from .components.node_editor import NodeEditor  # Пока нет
# from .components.format_dialog import FormatDialog  # Пока нет
# from .components.settings_dialog import SettingsDialog  # Пока нет


class MainWindow(QMainWindow):
    """
    Главное окно приложения Excel Micro DB (новый GUI).
    """

    def __init__(self, app_controller):
        """
        Инициализирует главное окно.

        Args:
            app_controller: Экземпляр AppController.
        """
        super().__init__()
        self.app_controller = app_controller
        self._setup_ui()
        self.setWindowTitle("Excel Micro DB (Новый GUI)")
        self.resize(1400, 900)

    def _setup_ui(self):
        """
        Настраивает пользовательский интерфейс.
        """
        # --- 1. Меню ---
        self._create_menu_bar()

        # --- 2. Панель инструментов (лента) ---
        self._create_tool_bar()

        # --- 3. Строка формул ---
        self._create_formula_bar()

        # --- 4. Центральный виджет (таблицы) ---
        self._create_central_widget()

        # --- 5. Вкладки листов ---
        # self._create_sheet_tabs() # Интегрировано в central_widget

        # --- 6. Док-панели ---
        self._create_dock_widgets()

        # --- 7. Строка состояния ---
        self._create_status_bar()

    def _create_menu_bar(self):
        """
        Создает строку меню.
        """
        menu_bar = self.menuBar()
        # Файл
        file_menu = menu_bar.addMenu("Файл")
        file_menu.addAction("Новый", self._on_new_project)
        file_menu.addAction("Открыть", self._on_open_project)
        file_menu.addAction("Сохранить", self._on_save_project)
        file_menu.addSeparator()
        file_menu.addAction("Выход", self.close)

        # Правка
        edit_menu = menu_bar.addMenu("Правка")
        edit_menu.addAction("Отменить", self._on_undo)
        edit_menu.addAction("Повторить", self._on_redo)
        edit_menu.addSeparator()
        edit_menu.addAction("Копировать", self._on_copy)
        edit_menu.addAction("Вставить", self._on_paste)

        # Вид
        view_menu = menu_bar.addMenu("Вид")
        view_menu.addAction("Обозреватель проекта", self._on_toggle_project_explorer)
        view_menu.addAction("Нодовый редактор", self._on_toggle_node_editor)

        # Инструменты
        tools_menu = menu_bar.addMenu("Инструменты")
        tools_menu.addAction("Настройки", self._on_settings)

        # Справка
        help_menu = menu_bar.addMenu("Справка")
        help_menu.addAction("О программе", self._on_about)

    def _create_tool_bar(self):
        """
        Создает панель инструментов (упрощённая лента).
        """
        toolbar = self.addToolBar("Основная")
        toolbar.addAction("Новый", self._on_new_project)
        toolbar.addAction("Открыть", self._on_open_project)
        toolbar.addAction("Сохранить", self._on_save_project)
        toolbar.addSeparator()
        toolbar.addAction("Отменить", self._on_undo)
        toolbar.addAction("Повторить", self._on_redo)
        toolbar.addSeparator()
        toolbar.addAction("Копировать", self._on_copy)
        toolbar.addAction("Вставить", self._on_paste)

    def _create_formula_bar(self):
        """
        Создает строку формул над центральным виджетом.
        """
        # Создаем виджет для строки формул
        formula_widget = QFrame()
        formula_layout = QHBoxLayout(formula_widget)
        formula_layout.setContentsMargins(0, 0, 0, 0)  # Убираем отступы

        # Метка для адреса ячейки
        self.cell_address_label = QLabel("A1")
        self.cell_address_label.setFixedWidth(60)
        self.cell_address_label.setAlignment(Qt.AlignCenter)
        self.cell_address_label.setStyleSheet("QLabel { background-color: #f0f0f0; border: 1px solid #a0a0a0; }")

        # Поле редактирования (сама строка формул)
        self.formula_line_edit = QLineEdit()
        self.formula_line_edit.setPlaceholderText("Введите формулу или значение")

        # Кнопка подтверждения (обычно не используется в Excel, но может быть полезна)
        confirm_button = QToolButton()
        confirm_button.setText("✓")
        confirm_button.clicked.connect(self._on_confirm_formula)

        # Кнопка отмены
        cancel_button = QToolButton()
        cancel_button.setText("✗")
        cancel_button.clicked.connect(self._on_cancel_formula)

        # Добавляем элементы в строку
        formula_layout.addWidget(self.cell_address_label)
        formula_layout.addWidget(self.formula_line_edit)
        formula_layout.addWidget(confirm_button)
        formula_layout.addWidget(cancel_button)

        # Добавляем строку формул в центральную область (над QTableView)
        # Для этого нужно будет немного изменить _create_central_widget
        # или вставить этот виджет в нужное место иерархии.
        # Пока добавим его как виджет в QMainWindow, но он должен быть над central_widget.
        # Это можно сделать, поместив central_widget в QSplitter или QVBoxLayout.
        # Но для простоты пока добавим его как отдельный виджет над центральным.
        # Более правильный способ - включить его в layout центрального виджета.
        # Центральный виджет будет QVBoxLayout, в котором сверху строка формул, а потом QTabWidget.
        # Это потребует небольшой перестройки _create_central_widget.
        # Пока оставлю так, и _create_central_widget будет учитывать строку формул.

        # Временно: устанавливаем как центральный виджет и добавляем в него строку формул и таблицу.
        # Это не идеально, но для начала сойдёт.
        central_widget_layout = QVBoxLayout()
        central_widget_layout.addWidget(formula_widget)

        # Центральный виджет для таблиц (пока просто QTabWidget)
        self.sheet_tabs = QTabWidget()
        central_widget_layout.addWidget(self.sheet_tabs)

        central_frame = QWidget()
        central_frame.setLayout(central_widget_layout)
        self.setCentralWidget(central_frame)

        # Теперь строка формул и вкладки находятся в центральном виджете.
        # self.setCentralWidget(self.sheet_tabs) # Закомментировано, так как теперь у нас QVBoxLayout в центральном виджете

    def _create_central_widget(self):
        """
        Создает центральный виджет с вкладками листов и таблицами.
        Учитывает строку формул, созданную в _create_formula_bar.
        """
        # Уже частично реализовано в _create_formula_bar через QVBoxLayout.
        # self.sheet_tabs уже создан там.
        # Здесь можно добавить SheetEditor в вкладки, когда он будет готов.
        # Пока добавим пустую вкладку для демонстрации.
        # dummy_table = QTableView()
        # self.sheet_tabs.addTab(dummy_table, "Лист1")

    def _create_dock_widgets(self):
        """
        Создает док-панели (обозреватель проекта, нодовый редактор).
        """
        # Обозреватель проекта (слева)
        project_explorer_dock = QDockWidget("Обозреватель проекта", self)
        project_explorer_widget = QWidget()  # Заменить на реальный виджет
        project_explorer_dock.setWidget(project_explorer_widget)
        project_explorer_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.LeftDockWidgetArea, project_explorer_dock)

        # Нодовый редактор (справа)
        node_editor_dock = QDockWidget("Нодовый редактор", self)
        node_editor_widget = QWidget()  # Заменить на реальный виджет (NodeGraphQt)
        node_editor_dock.setWidget(node_editor_widget)
        node_editor_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, node_editor_dock)

    def _create_status_bar(self):
        """
        Создает строку состояния.
        """
        status_bar = self.statusBar()
        status_bar.showMessage("Готов")

    # --- Заглушки для действий ---
    def _on_new_project(self):
        print("Новый проект")

    def _on_open_project(self):
        print("Открыть проект")

    def _on_save_project(self):
        print("Сохранить проект")

    def _on_undo(self):
        print("Отменить")

    def _on_redo(self):
        print("Повторить")

    def _on_copy(self):
        print("Копировать")

    def _on_paste(self):
        print("Вставить")

    def _on_toggle_project_explorer(self):
        print("Переключить обозреватель проекта")

    def _on_toggle_node_editor(self):
        print("Переключить нодовый редактор")

    def _on_settings(self):
        print("Настройки")

    def _on_about(self):
        print("О программе")

    def _on_confirm_formula(self):
        print(f"Подтвердить формулу: {self.formula_line_edit.text()}")

    def _on_cancel_formula(self):
        self.formula_line_edit.clear()
        print("Отменить редактирование")
    # --- Конец заглушек ---
