# demo_gui.py
"""
Демонстрационное приложение для сравнения стилей PySide6 и qt-material.
"""

import sys
from pathlib import Path
from typing import Optional

# Добавляем корень проекта в sys.path, чтобы можно было импортировать модули
project_root = Path(__file__).parent.resolve()
sys.path.insert(0, str(project_root))

# --- ИСПРАВЛЕНО: Импорт QtCore для доступа к атрибутам Qt ---
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QComboBox, QSpinBox, QSlider,
    QCheckBox, QRadioButton, QGroupBox, QMenuBar, QMenu, QToolBar, QStatusBar,
    QTabWidget, QTextEdit, QListWidget, QTreeWidget, QTreeWidgetItem
)
# --- КОНЕЦ ИСПРАВЛЕНИЯ ---

import logging

# Настройка логирования
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DemoMainWindow(QMainWindow):
    """Главное окно для демонстрации стилей."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Демо стилей PySide6")
        self.resize(800, 600)

        # --- Меню ---
        menubar = self.menuBar()
        file_menu = menubar.addMenu("&Файл")
        # --- ИСПРАВЛЕНО: Явное создание QAction ---
        file_menu.addAction(QAction("Новый", self))
        file_menu.addAction(QAction("Открыть", self))
        file_menu.addAction(QAction("Сохранить", self))
        # --- КОНЕЦ ИСПРАВЛЕНИЯ ---
        file_menu.addSeparator()
        file_menu.addAction(QAction("Выход", self))

        edit_menu = menubar.addMenu("&Правка")
        edit_menu.addAction(QAction("Отменить", self))
        edit_menu.addAction(QAction("Повторить", self))

        # --- Панель инструментов ---
        toolbar = self.addToolBar("Основная")
        # --- ИСПРАВЛЕНО: Явное создание QAction ---
        toolbar.addAction(QAction("Новый", self))
        toolbar.addAction(QAction("Открыть", self))
        toolbar.addAction(QAction("Сохранить", self))
        # --- КОНЕЦ ИСПРАВЛЕНИЯ ---

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
        # --- ИСПРАВЛЕНО: QtCore.Qt.Horizontal ---
        slider = QSlider(QtCore.Qt.Horizontal)
        # --- КОНЕЦ ИСПРАВЛЕНИЯ ---
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


# --- ИСПРАВЛЕНО: Аннотации типов ---
def run_demo(style_name: Optional[str] = "Fusion", material_theme: Optional[str] = None):
# --- КОНЕЦ ИСПРАВЛЕНИЯ ---
    """Запускает демонстрационное приложение с указанным стилем."""
    logger.info(f"Запуск демо с 'style_name={style_name}', 'material_theme={material_theme}'...")
    app = QApplication(sys.argv)
    app.setApplicationName("DemoGUI")

    # --- Установка стиля PySide6 ---
    if style_name:
        try:
            app.setStyle(style_name)
            logger.info(f"Установлен стиль PySide6: {style_name}")
        except Exception as e:
            logger.error(f"Ошибка при установке стиля PySide6 '{style_name}': {e}")
    else:
        logger.info("Стиль PySide6 не устанавливается.")

    # --- Установка темы qt-material ---
    if material_theme:
        try:
            import qt_material
            qt_material.apply_stylesheet(app, theme=material_theme)
            logger.info(f"Применена тема qt-material: {material_theme}")
        except ImportError:
            logger.error("Ошибка: Библиотека qt-material не установлена.")
            print("Ошибка: Библиотека qt-material не установлена.")
            return
        except Exception as e:
            logger.error(f"Ошибка при применении темы qt-material '{material_theme}': {e}", exc_info=True)
            print(f"Ошибка при применении темы qt-material '{material_theme}': {e}")
            return
    else:
        logger.info("Тема qt-material не применяется.")

    window = DemoMainWindow()
    window.show()
    logger.info("Демонстрационное окно открыто. Закройте его, чтобы продолжить.")
    print("Демонстрационное окно открыто. Закройте его, чтобы продолжить.")
    
    try:
        exit_code = app.exec()
        logger.info(f"Демонстрационное окно закрыто. Код завершения: {exit_code}")
        print(f"Демонстрационное окно закрыто. Код завершения: {exit_code}")
        return exit_code
    except Exception as e:
        logger.error(f"Ошибка при выполнении цикла событий приложения: {e}", exc_info=True)
        print(f"Ошибка при выполнении цикла событий приложения: {e}")
        return -1


if __name__ == "__main__":
    logger.info("--- Демонстрация стилей PySide6 ---")
    print("--- Демонстрация стилей PySide6 ---")
    
    # 1. Запуск с Fusion стилем
    logger.info("\n--- 1. Стиль Fusion ---")
    print("\n--- 1. Стиль Fusion ---")
    try:
        run_demo(style_name="Fusion", material_theme=None)
    except Exception as e:
        logger.error(f"Необработанная ошибка при запуске Fusion: {e}", exc_info=True)
        print(f"Необработанная ошибка при запуске Fusion: {e}")
    
    # 2. Запуск с темой qt-material (тёмная)
    logger.info("\n--- 2. Тема qt-material (dark_teal) ---")
    print("\n--- 2. Тема qt-material (dark_teal) ---")
    try:
        # Отключаем Fusion, чтобы qt-material работал корректно
        run_demo(style_name=None, material_theme='dark_teal.xml') 
    except Exception as e:
        logger.error(f"Необработанная ошибка при запуске qt-material (dark_teal): {e}", exc_info=True)
        print(f"Необработанная ошибка при запуске qt-material (dark_teal): {e}")
    
    # 3. Запуск с темой qt-material (светлая)
    logger.info("\n--- 3. Тема qt-material (light_blue) ---")
    print("\n--- 3. Тема qt-material (light_blue) ---")
    try:
        # Отключаем Fusion, чтобы qt-material работал корректно
        run_demo(style_name=None, material_theme='light_blue.xml') 
    except Exception as e:
        logger.error(f"Необработанная ошибка при запуске qt-material (light_blue): {e}", exc_info=True)
        print(f"Необработанная ошибка при запуске qt-material (light_blue): {e}")

    logger.info("\n--- Демонстрация завершена ---")
    print("\n--- Демонстрация завершена ---")
