# gui.py
"""
Точка входа для графического интерфейса Excel Micro DB.
Упрощенная версия без концепции "проекта" - работает напрямую с Excel-файлами.
"""
import sys
import os
from pathlib import Path

# --- Добавление корня проекта в sys.path ---
project_root = Path(__file__).parent.resolve()
sys.path.insert(0, str(project_root))
# -------------------------------------------

# Импорт Qt
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import Qt, QCoreApplication
import logging

# Импорт упрощенного главного окна
from backend.constructor.widgets.simple_gui.gui_app import main as simple_gui_main

# Импорт логгера
from backend.utils.logger import get_logger

# Получаем логгер для этого модуля
logger = get_logger(__name__)

def main():
    """Главная функция для запуска GUI."""
    logger.info("Запуск упрощенного графического интерфейса Excel Micro DB")

    # Создание экземпляра QApplication
    app = QApplication(sys.argv)
    
    # Установка имени приложения
    QCoreApplication.setApplicationName("Excel Micro DB")
    QCoreApplication.setOrganizationName("ExcelMicroDB")
    QCoreApplication.setApplicationVersion("0.1.0")

    try:
        # Запускаем упрощенный GUI
        return simple_gui_main()

    except Exception as e:
        logger.critical(f"Критическая ошибка при запуске GUI: {e}", exc_info=True)
        QMessageBox.critical(
            None, 
            "Критическая ошибка", 
            f"Произошла критическая ошибка при запуске приложения:\n{e}\n\nПриложение будет закрыто."
        )
        sys.exit(1)

if __name__ == "__main__":
    main()