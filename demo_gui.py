# demo_gui.py
"""
Демонстрационное приложение для сравнения стилей PySide6 и qt-material.
"""

import sys
from pathlib import Path

# Добавляем корень проекта в sys.path, чтобы можно было импортировать модули
project_root = Path(__file__).parent.resolve()
sys.path.insert(0, str(project_root))

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QComboBox, QSpinBox, QSlider,
    QCheckBox, QRadioButton, QGroupBox, QMenuBar, QMenu, QToolBar, QStatusBar,
    QTabWidget, QTextEdit, QListWidget, QTreeWidget, QTreeWidgetItem
)
from PySide6.QtCore import Qt, Slot


class DemoMainWindow(QMainWindow):
    """Главное окно для демонстрации стилей."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Демо стилей PySide6")
        self.resize(800, 600)

        # --- Меню ---
        menubar = self.menuBar()
        file_menu = menubar.addMenu("&Файл")
        file_menu.addAction("Новый")
        file_menu.addAction("Открыть")
        file_menu.addAction("Сохранить")
        file_menu.addSeparator()
        file_menu.addAction("Выход")

        edit_menu = menubar.addMenu("&Правка")
        edit_menu.addAction("Отменить")
        edit_menu.addAction("Повторить")

        # --- Панель инструментов ---
        toolbar = self.addToolBar("Основная")
        toolbar.addAction("Новый")
        toolbar.addAction("Открыть")
        toolbar.addAction("Сохранить")

        # --- Центральный виджет ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- Левая панель (виджеты) ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        # Метка
        label = QLabel("Это метка (QLabel)")
        left_layout.addWidget(label)

        # Кнопка
        button = QPushButton("Это кнопка (QPushButton)")
        left_layout.addWidget(button)

        # Комбобокс
        combo = QComboBox()
        combo.addItems(["Элемент 1", "Элемент 2", "Элемент 3"])
        left_layout.addWidget(combo)

        # Спинбокс
        spinbox = QSpinBox()
        spinbox.setValue(42)
        left_layout.addWidget(spinbox)

        # Слайдер
        slider = QSlider(Qt.Horizontal)
        slider.setValue(50)
        left_layout.addWidget(slider)

        # Чекбокс
        checkbox = QCheckBox("Это чекбокс (QCheckBox)")
        left_layout.addWidget(checkbox)

        # Радиокнопки
        radio_group = QGroupBox("Радиокнопки (QRadioButton)")
        radio_layout = QVBoxLayout(radio_group)
        radio1 = QRadioButton("Вариант 1")
        radio2 = QRadioButton("Вариант 2")
        radio_layout.addWidget(radio1)
        radio_layout.addWidget(radio2)
        left_layout.addWidget(radio_group)

        # Список
        list_widget = QListWidget()
        list_widget.addItems(["Элемент списка 1", "Элемент списка 2", "Элемент списка 3"])
        left_layout.addWidget(list_widget)

        left_layout.addStretch() # Растягиваем пустое пространство

        # --- Правая панель (вкладки) ---
        tab_widget = QTabWidget()
        
        # Вкладка 1: Текстовый редактор
        text_edit = QTextEdit()
        text_edit.setPlainText("Это QTextEdit.\nМногострочное текстовое поле.")
        tab_widget.addTab(text_edit, "Текст")

        # Вкладка 2: Дерево
        tree_widget = QTreeWidget()
        tree_widget.setHeaderLabels(["Колонка 1", "Колонка 2"])
        root_item = QTreeWidgetItem(["Корневой элемент"])
        child_item1 = QTreeWidgetItem(["Дочерний 1", "Данные 1"])
        child_item2 = QTreeWidgetItem(["Дочерний 2", "Данные 2"])
        root_item.addChild(child_item1)
        root_item.addChild(child_item2)
        tree_widget.addTopLevelItem(root_item)
        tree_widget.expandAll()
        tab_widget.addTab(tree_widget, "Дерево")

        # Добавляем панели в основной layout
        main_layout.addWidget(left_widget, 1) # 1/3 ширины
        main_layout.addWidget(tab_widget, 2) # 2/3 ширины

        # --- Строка состояния ---
        self.statusBar().showMessage("Готов")


def run_demo(style_name: str = "Fusion", material_theme: str = None):
    """Запускает демонстрационное приложение с указанным стилем."""
    print(f"Запуск демо с '{style_name}' стилем...")
    app = QApplication(sys.argv)
    app.setApplicationName("DemoGUI")

    # --- Установка стиля ---
    if style_name:
        app.setStyle(style_name)
        print(f"Установлен стиль: {style_name}")

    # --- Установка темы qt-material ---
    if material_theme:
        try:
            import qt_material
            qt_material.apply_stylesheet(app, theme=material_theme)
            print(f"Применена тема qt-material: {material_theme}")
        except ImportError:
            print("Ошибка: Библиотека qt-material не установлена.")
            return
        except Exception as e:
            print(f"Ошибка при применении темы qt-material: {e}")
            return

    window = DemoMainWindow()
    window.show()
    print("Демонстрационное окно открыто. Закройте его, чтобы продолжить.")
    app.exec()
    print("Демонстрационное окно закрыто.")


if __name__ == "__main__":
    print("--- Демонстрация стилей PySide6 ---")
    
    # 1. Запуск с Fusion стилем
    print("\n--- 1. Стиль Fusion ---")
    run_demo(style_name="Fusion", material_theme=None)
    
    # 2. Запуск с темой qt-material (тёмная)
    print("\n--- 2. Тема qt-material (dark_teal) ---")
    run_demo(style_name=None, material_theme='dark_teal.xml') # Отключаем Fusion, чтобы qt-material работал корректно
    
    # 3. Запуск с темой qt-material (светлая)
    print("\n--- 3. Тема qt-material (light_blue) ---")
    run_demo(style_name=None, material_theme='light_blue.xml') # Отключаем Fusion, чтобы qt-material работал корректно

    print("\n--- Демонстрация завершена ---")
