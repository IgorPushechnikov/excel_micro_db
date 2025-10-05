# run_gui.py
"""
Точка входа для запуска GUI приложения Excel Micro DB.
"""

import sys
import os
import logging
from pathlib import Path

# Импортируем QApplication и MainWindow
from PySide6.QtWidgets import QApplication as QtApp
from backend.constructor.widgets.new_gui.main_window import MainWindow

# Импортируем логгер
# from backend.utils.logger import setup_logging, get_logger # <-- Убран setup_logging из импорта
from backend.utils.logger import get_logger, setup_logger # <-- Добавлен setup_logger


def main():
    """
    Основная функция запуска GUI.
    """
    # Настройка логирования
    # Путь к лог-файлу GUI можно сделать отдельным, например, в папке проекта
    # или использовать существующую систему из backend.utils.logger
    # Вызываем setup_logger для инициализации системы логирования
    # Пока что вызываем без аргументов, но можно добавить путь к логу GUI
    setup_logger()
    logger = get_logger(__name__)

    logger.info("Запуск GUI приложения Excel Micro DB...")

    # Инициализация QApplication
    app = QtApp(sys.argv)
    app.setApplicationName("Excel Micro DB GUI")
    app.setApplicationVersion("0.1.0")

    # Создание и отображение главного окна
    main_window = MainWindow()
    main_window.show()

    logger.info("GUI приложение запущено.")

    # Запуск цикла событий Qt
    exit_code = app.exec()
    logger.info(f"GUI приложение завершено с кодом: {exit_code}")

    sys.exit(exit_code)

if __name__ == "__main__":
    main()
