# src/constructor/gui_app.py
"""
Точка входа для графического интерфейса Excel Micro DB.
"""

import sys
import os
from pathlib import Path

# Добавляем корень проекта в sys.path, чтобы можно было импортировать модули
project_root = Path(__file__).parent.parent.parent.resolve()
sys.path.insert(0, str(project_root))

from PySide6.QtWidgets import QApplication
# from PySide6.QtCore import Qt # Пока не используем

# Импортируем AppController
from src.core.app_controller import create_app_controller

# Импортируем GUI Controller
from src.constructor.gui_controller import GUIController


def main():
    """Основная функция запуска GUI."""
    print("Запуск графического интерфейса Excel Micro DB...")

    # 1. Создаём QApplication
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Micro DB")
    app.setApplicationVersion("0.1.0")
    # app.setStyle("Fusion") # Можно включить позже

    # 2. Создаём AppController
    # Передаём путь к рабочей области, если нужно, или None для дефолтного
    # project_path = Path.home() / "excel_micro_db_workspace" # Пример
    project_path = None
    app_controller = create_app_controller(project_path=project_path)
    if not app_controller.initialize():
         print("Ошибка инициализации AppController.")
         return -1

    # 3. Создаём и запускаем GUIController
    try:
        gui_controller = GUIController(app_controller)
        gui_controller.run() # Этот метод будет показывать MainWindow и запускать app.exec()
        return_code = app.exec()
        print(f"Приложение завершено с кодом: {return_code}")
        return return_code
    except Exception as e:
        print(f"Критическая ошибка в GUI: {e}")
        # import traceback
        # traceback.print_exc() # Для отладки
        return -2


if __name__ == "__main__":
    # Для корректного отображения SVG и других ресурсов
    # QCoreApplication.setAttribute(Qt.AA_UseSoftwareOpenGL)
    # QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling) # Обычно включено по умолчанию
    # QCoreApplication.setAttribute(Qt.AA_UseHighDpiPixmaps) # Обычно включено по умолчанию
    sys.exit(main())
