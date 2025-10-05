# run_new_gui.py
"""
Точка входа для запуска нового GUI приложения Excel Micro DB.
Следует новому дизайну: таблица - центральный элемент.
"""

import sys
import os
import logging
from pathlib import Path

# --- Добавление корня проекта в sys.path ---
# Это необходимо для корректного импорта модулей из backend
project_root = Path(__file__).parent.resolve()
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))
# -------------------------------------------

# Импортируем QApplication
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QCoreApplication, Qt
from PySide6.QtGui import QStyleFactory

# Импортируем новое главное окно
# Убедимся, что импортируем из правильного модуля
# Если main_window_new.py был переименован в main_window.py, путь будет другой
# from backend.constructor.widgets.new_gui.main_window_new import MainWindowNew
# Предположим, что мы будем использовать обновлённый main_window.py
from backend.constructor.widgets.new_gui.main_window import MainWindow

# Импортируем логгер
from backend.utils.logger import get_logger, setup_logger

logger = get_logger(__name__)

def main():
    """
    Основная функция запуска нового GUI.
    """
    # Настройка логирования
    # setup_logger() вызывается внутри MainWindow, но можно и здесь для раннего лога
    setup_logger() 
    logger.info("Запуск НОВОГО GUI приложения Excel Micro DB...")

    try:
        # Создание экземпляра QApplication
        app = QApplication(sys.argv)
        
        # Установка имени и версии приложения
        QCoreApplication.setApplicationName("Excel Micro DB New GUI")
        QCoreApplication.setOrganizationName("ExcelMicroDB")
        QCoreApplication.setApplicationVersion("0.2.0")
        
        # Установка стиля (опционально, для лучшего внешнего вида)
        app.setStyle(QStyleFactory.create("Fusion"))

        # Создание и отображение главного окна
        # Используем обновлённый MainWindow, который теперь следует новому дизайну
        window = MainWindow()
        window.show()

        logger.info("НОВОЕ GUI приложение запущено.")

        # Запуск цикла событий Qt
        exit_code = app.exec()
        logger.info(f"НОВОЕ GUI приложение завершено с кодом: {exit_code}")

        sys.exit(exit_code)

    except Exception as e:
        logger.critical(f"Критическая ошибка при запуске НОВОГО GUI: {e}", exc_info=True)
        # Используем стандартный print, так как QApplication может не быть создан
        print(f"[КРИТИЧЕСКАЯ ОШИБКА] Произошла критическая ошибка при запуске приложения:\n{e}\n\nПриложение будет закрыто.", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
