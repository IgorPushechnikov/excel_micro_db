# src/constructor_new/gui_controller.py
"""
Модуль для центрального контроллера графического интерфейса (новый GUI).
Управляет жизненным циклом GUI и координирует взаимодействие
между AppController, MainWindow и другими компонентами.
"""

import logging
from typing import Optional

from PySide6.QtWidgets import QMessageBox, QApplication
from PySide6.QtCore import QObject, Slot

# Импортируем AppController
from src.core.app_controller import AppController

# Импортируем основное окно
from .main_window import MainWindow

# Получаем логгер
logger = logging.getLogger(__name__)


class GUIController(QObject):
    """
    Центральный контроллер графического интерфейса (новый GUI).
    """

    def __init__(self, app: QApplication, app_controller: AppController):
        """
        Инициализирует контроллер GUI.

        Args:
            app (QApplication): Экземпляр QApplication.
            app_controller (AppController): Экземпляр основного контроллера приложения.
        """
        super().__init__()
        self.app: QApplication = app
        self.app_controller: AppController = app_controller
        self.main_window: Optional[MainWindow] = None
        logger.debug("GUIController (новый) инициализирован.")

    def run(self) -> int:
        """
        Запускает графический интерфейс и возвращает код завершения.

        Returns:
            int: Код завершения приложения (обычно 0 для успеха).
        """
        logger.info("Запуск нового графического интерфейса...")
        try:
            # 1. Создаём главное окно
            self.main_window = MainWindow(self.app_controller)

            # 2. Подключаем сигналы/слоты между MainWindow и GUIController (если нужно)
            # Пример: сигнал закрытия главного окна может вызывать shutdown
            # self.main_window.some_signal.connect(self.some_slot)

            # 3. Показываем главное окно
            self.main_window.show()
            logger.info("Главное окно нового GUI отображено.")

            # 4. Запускаем цикл событий Qt
            logger.info("Запуск цикла событий приложения...")
            return self.app.exec()

        except Exception as e:
            logger.error(f"Ошибка при запуске нового GUI: {e}", exc_info=True)
            # Можно показать критическое сообщение об ошибке
            QMessageBox.critical(
                None,
                "Критическая ошибка GUI",
                f"Не удалось запустить новый графический интерфейс:\\n{e}"
            )
            return -1 # Код ошибки

    @Slot()
    def shutdown(self):
        """
        Корректно завершает работу GUI.
        """
        logger.info("Завершение работы GUIController (новый)...")
        # Здесь можно выполнить очистку, сохранить состояние и т.д.
        # Пока просто логируем.
        if self.app_controller:
            self.app_controller.shutdown()
        logger.info("GUIController (новый) завершил работу.")
