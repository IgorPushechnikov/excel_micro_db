# backend/constructor/widgets/simple_gui/gui_app.py
"""
Точка входа для упрощённого графического интерфейса Excel Micro DB.
"""
import sys
import os
from pathlib import Path

# --- Добавление корня проекта в sys.path ---
project_root = Path(__file__).parent.parent.parent.parent.resolve()
sys.path.insert(0, str(project_root))
# -------------------------------------------

# Импорт Qt
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import Qt, QCoreApplication
import logging

# Импорт упрощенного главного окна
from backend.constructor.widgets.simple_gui.main_window import SimpleMainWindow

# Импорт логгера
from backend.utils.logger import get_logger

# Получаем логгер для этого модуля
logger = get_logger(__name__)

def main():
    """Главная функция для запуска упрощенного GUI."""
    logger.info("Запуск упрощенного графического интерфейса Excel Micro DB")

    # Проверяем, существует ли уже экземпляр QApplication
    app = QApplication.instance()
    if app is None:
        # Создание экземпляра QApplication только если его нет
        app = QApplication(sys.argv)
        logger.debug("Создан новый экземпляр QApplication")
    else:
        logger.debug("Используется существующий экземпляр QApplication")
    
    # Установка имени приложения
    QCoreApplication.setApplicationName("Excel Micro DB")
    QCoreApplication.setOrganizationName("ExcelMicroDB")
    QCoreApplication.setApplicationVersion("0.1.0")

    try:
        window = SimpleMainWindow()
        window.show()
        logger.info("Упрощенное главное окно отображено")
        
        exit_code = app.exec()
        logger.info(f"Цикл событий завершен с кодом {exit_code}")
        return exit_code

    except Exception as e:
        logger.critical(f"Критическая ошибка при запуске упрощенного GUI: {e}", exc_info=True)
        QMessageBox.critical(
            None, 
            "Критическая ошибка", 
            f"Произошла критическая ошибка:\n{e}\n\nПриложение будет закрыто."
        )
        return 1

if __name__ == "__main__":
    sys.exit(main())