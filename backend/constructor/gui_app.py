# backend/constructor/gui_app.py
"""
Модуль для запуска графического интерфейса Excel Micro DB.
Это точка входа, импортируемая main.py при вызове с аргументом --gui.
"""

import sys
from pathlib import Path

# Импорты PySide6
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QCoreApplication

# Импорт нашего главного окна
from .main_window import MainWindow  # Используем относительный импорт, так как gui_app.py будет в backend/constructor/

def main():
    """
    Основная функция для запуска GUI.
    Создает QApplication, устанавливает параметры и показывает MainWindow.
    """
    # Создание экземпляра QApplication
    app = QApplication(sys.argv)

    # (Опционально) Установка имени приложения для стилей/настроек ОС
    QCoreApplication.setApplicationName("Excel Micro DB")
    QCoreApplication.setOrganizationName("YourOrgOrName") # Замените на ваше
    QCoreApplication.setApplicationVersion("0.1.0") # Замените на актуальную версию

    # Создание и отображение главного окна
    window = MainWindow()
    window.show()

    # Запуск цикла обработки событий приложения
    # exec_ возвращает код завершения, который мы передаем в sys.exit()
    exit_code = app.exec()
    return exit_code

if __name__ == "__main__":
    # Если этот файл запускается напрямую (python gui_app.py),
    # вызываем main() и передаем его результат в sys.exit()
    sys.exit(main())