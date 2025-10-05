# backend/constructor/widgets/new_gui/import_worker.py
"""
Вспомогательный класс для выполнения импорта в отдельном потоке.
"""

import logging
from pathlib import Path
from typing import Optional, Callable # <-- Добавлен Callable
from PySide6.QtCore import QThread, Signal

# Импортируем AppController
from backend.core.app_controller import create_app_controller
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ImportWorker(QThread):
    """
    Рабочий поток для выполнения импорта данных через AppController.
    """
    finished = Signal(bool, str)  # (успех/ошибка, сообщение)
    progress = Signal(int, str)   # (значение, сообщение) - если AppController будет передавать прогресс

    def __init__(self, app_controller, file_path, import_type, import_mode, progress_callback: Optional[Callable[[int, str], None]] = None): # <-- ИЗМЕНЕНО
        """
        Инициализирует рабочий поток импорта.

        Args:
            app_controller: Экземпляр AppController.
            file_path (str): Путь к импортируемому файлу.
            import_type (str): Тип импорта.
            import_mode (str): Режим импорта.
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса. # <-- ДОБАВЛЕНО
        """
        super().__init__()
        self.app_controller = app_controller
        self.file_path = file_path
        self.import_type = import_type
        self.import_mode = import_mode
        # --- НОВОЕ: Получаем путь к БД из app_controller ---
        self.db_path = app_controller.project_db_path
        # --- КОНЕЦ НОВОГО ---
        self.progress_callback = progress_callback # <-- НОВОЕ

    def run(self):
        """
        Запускает импорт в отдельном потоке.
        """
        try:
            logger.info(f"Начало импорта (тип: {self.import_type}, режим: {self.import_mode}) для файла {self.file_path} в потоке {id(QThread.currentThread())}")

            # --- НОВОЕ: Создаём функцию-обёртку для прогресса ---
            def internal_progress_callback(value, message):
                # Эмитим сигнал для внутреннего использования (например, в MainWindow)
                self.progress.emit(value, message)
                # Если передан внешний callback, вызываем его
                if self.progress_callback:
                    self.progress_callback(value, message)
            # --- КОНЕЦ НОВОГО ---

            # Определяем метод AppController на основе типа и режима
            method_name = f"import_{self.import_type}_from_excel"
            if self.import_mode != 'all' and self.import_mode != 'fast':
                 method_name += f"_{self.import_mode}"

            method = getattr(self.app_controller, method_name, None)
            if method is None:
                raise AttributeError(f"AppController не имеет метода {method_name}")

            # --- ИЗМЕНЕНО: Передаем db_path и progress_callback ---
            # success = method(self.file_path, db_path=self.db_path)
            success = method(self.file_path, db_path=self.db_path, progress_callback=internal_progress_callback) # <-- ИЗМЕНЕНО
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            logger.info(f"Импорт (тип: {self.import_type}, режим: {self.import_mode}) для файла {self.file_path} завершён в потоке {id(QThread.currentThread())}.")

            # Отправляем результат
            self.finished.emit(success, f"Импорт ({self.import_type}, {self.import_mode}) {'успешен' if success else 'неудачен'}.")

        except Exception as e:
            logger.error(f"Ошибка в потоке импорта для файла {self.file_path} (тип: {self.import_type}, режим: {self.import_mode}): {e}", exc_info=True)
            # Отправляем ошибку
            self.finished.emit(False, f"Ошибка импорта: {e}")
