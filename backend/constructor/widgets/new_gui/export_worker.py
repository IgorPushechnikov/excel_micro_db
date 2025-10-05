# backend/constructor/widgets/new_gui/export_worker.py
"""
Вспомогательный класс для выполнения экспорта в отдельном потоке.
"""

import logging
from pathlib import Path
from typing import Optional, Callable
from PySide6.QtCore import QThread, Signal

# Импортируем AppController
from backend.core.app_controller import create_app_controller
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ExportWorker(QThread):
    """
    Рабочий поток для выполнения экспорта данных через AppController.
    """
    finished = Signal(bool, str)  # (успех/ошибка, сообщение)
    progress = Signal(int, str)   # (значение, сообщение) - если AppController будет передавать прогресс

    def __init__(self, app_controller, output_path, progress_callback: Optional[Callable[[int, str], None]] = None):
        """
        Инициализирует рабочий поток экспорта.

        Args:
            app_controller: Экземпляр AppController.
            output_path (str): Путь к файлу, в который будет экспортировано.
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
        """
        super().__init__()
        self.app_controller = app_controller
        self.output_path = output_path
        # --- НОВОЕ: Получаем путь к БД из app_controller ---
        self.db_path = app_controller.project_db_path
        # --- КОНЕЦ НОВОГО ---
        self.progress_callback = progress_callback

    def run(self):
        """
        Запускает экспорт в отдельном потоке.
        """
        try:
            logger.info(f"Начало экспорта в файл {self.output_path} в потоке {id(QThread.currentThread())}")

            # --- НОВОЕ: Создаём функцию-обёртку для прогресса ---
            def internal_progress_callback(value, message):
                # Эмитим сигнал для внутреннего использования (например, в MainWindow)
                self.progress.emit(value, message)
                # Если передан внешний callback, вызываем его
                if self.progress_callback:
                    self.progress_callback(value, message)
            # --- КОНЕЦ НОВОГО ---

            # --- ИЗМЕНЕНО: Логика вызова метода AppController для экспорта ---
            # Предположим, AppController имеет метод export_to_excel
            method_name = "export_to_excel" # <-- Пример имени метода
            method = getattr(self.app_controller, method_name, None)
            if method is None:
                raise AttributeError(f"AppController не имеет метода {method_name}")

            # Вызываем метод экспорта
            # --- ИЗМЕНЕНО: Передаем output_path и db_path, а также progress_callback ---
            success = method(self.output_path, db_path=self.db_path, progress_callback=internal_progress_callback)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            logger.info(f"Экспорт в файл {self.output_path} завершён в потоке {id(QThread.currentThread())}.")

            # Отправляем результат
            self.finished.emit(success, f"Экспорт в {self.output_path} {'успешен' if success else 'неудачен'}.")

        except Exception as e:
            logger.error(f"Ошибка в потоке экспорта для файла {self.output_path}: {e}", exc_info=True)
            # Отправляем ошибку
            self.finished.emit(False, f"Ошибка экспорта: {e}")