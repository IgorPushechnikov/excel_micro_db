# src/constructor_new/gui_app.py
"""
Точка входа для нового графического интерфейса Excel Micro DB.
"""

import sys
from pathlib import Path

# Добавляем корень проекта в путь поиска модулей, если нужно (обычно не обязательно, если установлен как пакет)
# project_root = Path(__file__).parent.parent.parent.resolve()
# sys.path.insert(0, str(project_root))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QSettings

# Импортируем AppController
from src.core.app_controller import create_app_controller

# Импортируем основное окно и контроллер
from .main_window import MainWindow
from .gui_controller import GUIController # Предполагается, что будет создан


def main():
    """
    Основная функция запуска GUI.
    """
    print("Запуск нового графического интерфейса Excel Micro DB...")

    # 1. Создаём QApplication
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Micro DB New GUI")
    app.setApplicationVersion("0.1.0")

    # --- НОВОЕ: Загрузка и применение темы ---
    settings = QSettings("ExcelMicroDB", "NewGUI") # Используем QSettings для хранения настроек
    theme_name = settings.value("theme", "default") # "default", "dark", "light", или имя .qss файла

    if theme_name and theme_name != "default":
        try:
            # Путь к файлу темы, например, в папке resources/themes/
            # theme_file_path = Path(__file__).parent.parent / "resources" / "themes" / f"{theme_name}.qss"
            # Для простоты, пока используем имя темы напрямую, если это стандартная тема qt-material
            # или путь к .qss файлу.
            # Пока оставим заглушку для QSS.
            theme_file_path = Path(f"themes/{theme_name}.qss") # Пример пути
            if theme_file_path.exists():
                with open(theme_file_path, 'r', encoding='utf-8') as f:
                    app.setStyleSheet(f.read())
                print(f"Применена тема из файла: {theme_file_path}")
            else:
                print(f"Файл темы {theme_file_path} не найден. Используется тема по умолчанию.")
        except Exception as e:
            print(f"Ошибка при загрузке темы {theme_name}: {e}")
    # --- КОНЕЦ НОВОГО ---
    # --- АЛЬТЕРНАТИВА: Использование qt-material ---
    # if theme_name == "dark":
    #     import qt_material
    #     qt_material.apply_stylesheet(app, theme='dark_teal.xml')
    # elif theme_name == "light":
    #     import qt_material
    #     qt_material.apply_stylesheet(app, theme='light_blue.xml')
    # --- КОНЕЦ АЛЬТЕРНАТИВЫ ---

    # 2. Создаём AppController
    # Передаём путь к рабочей области, если нужно, или None для дефолтного
    project_path = None # Или Path("путь/к/проекту") или через QSettings
    app_controller = create_app_controller(project_path=project_path)
    if not app_controller.initialize():
         print("Ошибка инициализации AppController.")
         return -1

    # 3. Создаём и запускаем GUIController
    try:
        # Передаём app_controller в GUIController
        gui_controller = GUIController(app, app_controller)
        # GUIController отвечает за создание MainWindow и запуск app.exec()
        return_code = gui_controller.run() # run теперь возвращает код
        print(f"Приложение завершено с кодом: {return_code}")
        return return_code
    except Exception as e:
        print(f"Критическая ошибка в GUI: {e}")
        import traceback
        traceback.print_exc() # Для отладки
        return -2


if __name__ == "__main__":
    # Для корректного отображения SVG и других ресурсов
    # QApplication.setAttribute(Qt.AA_UseSoftwareOpenGL)
    # QApplication.setAttribute(Qt.AA_EnableHighDpiScaling) # Обычно включено по умолчанию
    # QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps) # Обычно включено по умолчанию
    sys.exit(main())
