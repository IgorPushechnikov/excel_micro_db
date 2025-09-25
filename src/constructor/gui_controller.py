# src/constructor/gui_controller.py
"""
Модуль для центрального контроллера графического интерфейса.
Управляет жизненным циклом GUI и координирует взаимодействие
между AppController и виджетами UI.
"""

import logging
from typing import Optional

from PySide6.QtWidgets import QMessageBox
from PySide6.QtCore import QObject, Slot, Signal

# Импортируем AppController
from src.core.app_controller import AppController

# Импортируем виджеты
from src.constructor.widgets.main_window import MainWindow

# Получаем логгер
logger = logging.getLogger(__name__)


class GUIController(QObject):
    """
    Центральный контроллер графического интерфейса.
    """

    def __init__(self, app_controller: AppController):
        """
        Инициализирует контроллер GUI.

        Args:
            app_controller (AppController): Экземпляр основного контроллера приложения.
        """
        super().__init__()
        self.app_controller: AppController = app_controller
        self.main_window: Optional[MainWindow] = None
        logger.debug("GUIController инициализирован.")

    def run(self):
        """Запускает графический интерфейс."""
        logger.info("Запуск графического интерфейса...")
        try:
            # 1. Создаём главное окно
            self.main_window = MainWindow(self.app_controller)
            
            # 2. Подключаем сигналы/слоты между MainWindow и GUIController (если нужно)
            # Например, сигнал закрытия главного окна может вызывать shutdown
            # self.main_window.some_signal.connect(self.some_slot)

            # 3. Показываем главное окно
            self.main_window.show()
            logger.info("Главное окно GUI отображено.")

        except Exception as e:
            logger.error(f"Ошибка при запуске GUI: {e}", exc_info=True)
            # Можно показать критическое сообщение об ошибке
            QMessageBox.critical(
                None, 
                "Критическая ошибка GUI", 
                f"Не удалось запустить графический интерфейс:\n{e}"
            )
            raise # Повторно поднимаем исключение

    @Slot()
    def shutdown(self):
        """Корректно завершает работу GUI."""
        logger.info("Завершение работы GUIController...")
        # Здесь можно выполнить очистку, сохранить состояние и т.д.
        # Пока просто логируем.
        if self.app_controller:
            self.app_controller.shutdown()
        logger.info("GUIController завершил работу.")
