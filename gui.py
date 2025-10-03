# gui.py
"""
Точка входа для графического интерфейса Excel Micro DB.
"""
import sys
import os
from pathlib import Path

# --- Добавление корня проекта в sys.path ---
# Это необходимо, чтобы модули из backend/ были видны.
# sys.path.insert(0, str(Path(__file__).parent / "backend")) - УДАЛЕНО
# Вместо этого, добавим родительскую директорию проекта, чтобы импортировать как backend.constructor...
project_root = Path(__file__).parent.resolve()
sys.path.insert(0, str(project_root))
# -------------------------------------------

# Импорт Qt
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import Qt, QCoreApplication
import logging

# Импорт нашего главного окна
# ИСПРАВЛЕНО: Абсолютный импорт относительно корня проекта, где backend в sys.path
from backend.constructor.main_window import MainWindow

# Импорт логгера
# ИСПРАВЛЕНО: Абсолютный импорт относительно корня проекта, где backend в sys.path
from backend.utils.logger import get_logger

# Получаем логгер для этого модуля
logger = get_logger(__name__)

def main():
    """Главная функция для запуска GUI."""
    logger.info("Запуск графического интерфейса Excel Micro DB")

    # Создание экземпляра QApplication
    # Это обязательный компонент для любого Qt-приложения
    app = QApplication(sys.argv)
    
    # (Опционально) Установка имени приложения для стилей/настроек ОС
    QCoreApplication.setApplicationName("Excel Micro DB")
    QCoreApplication.setOrganizationName("YourOrgOrName") # Замените на ваше
    QCoreApplication.setApplicationVersion("0.1.0") # Замените на актуальную версию

    try:
        # Создание и отображение главного окна
        logger.debug("Создание экземпляра MainWindow")
        window = MainWindow()
        window.show()
        logger.info("Главное окно отображено")

        # Запуск цикла обработки событий приложения
        logger.debug("Запуск цикла событий QApplication")
        exit_code = app.exec()
        logger.info(f"Цикл событий QApplication завершен с кодом {exit_code}")
        sys.exit(exit_code)

    except Exception as e:
        logger.critical(f"Критическая ошибка при запуске GUI: {e}", exc_info=True)
        # Показ критической ошибки пользователю, если GUI уже инициализирован
        QMessageBox.critical(
            None, 
            "Критическая ошибка", 
            f"Произошла критическая ошибка при запуске приложения:\n{e}\n\nПриложение будет закрыто."
        )
        sys.exit(1)

if __name__ == "__main__":
    main()
