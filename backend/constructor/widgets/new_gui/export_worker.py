# backend/constructor/widgets/new_gui/import_worker.py
"""
Вспомогательный класс для выполнения импорта в отдельном потоке.
"""

import logging
from pathlib import Path
from typing import Optional, Callable
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

    def __init__(self, app_controller, file_path, import_mode_key, progress_callback: Optional[Callable[[int, str], None]] = None):
        """
        Инициализирует рабочий поток импорта.

        Args:
            app_controller: Экземпляр AppController.
            file_path (str): Путь к импортируемому файлу.
            import_mode_key (str): Ключ объединённого режима импорта (например, 'all_openpyxl').
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
        """
        super().__init__()
        self.app_controller = app_controller
        self.file_path = file_path
        self.import_mode_key = import_mode_key # <-- Изменено: один ключ
        # --- НОВОЕ: Получаем путь к БД из app_controller ---
        self.db_path = app_controller.project_db_path
        # --- КОНЕЦ НОВОГО ---
        self.progress_callback = progress_callback

    def run(self):
        """
        Запускает импорт в отдельном потоке.
        """
        try:
            logger.info(f"Начало импорта (режим: {self.import_mode_key}) для файла {self.file_path} в потоке {id(QThread.currentThread())}")

            # --- НОВОЕ: Создаём функцию-обёртку для прогресса ---
            def internal_progress_callback(value, message):
                # Эмитим сигнал для внутреннего использования (например, в MainWindow)
                self.progress.emit(value, message)
                # Если передан внешний callback, вызываем его
                if self.progress_callback:
                    self.progress_callback(value, message)
            # --- КОНЕЦ НОВОГО ---

            # --- ИЗМЕНЕНО: Логика выбора метода на основе import_mode_key ---
            # Предположим, AppController имеет метод import_by_mode_key
            # или мы можем явно определить методы для каждого ключа здесь.
            # Пока используем условный вызов import_by_mode_key.
            # В будущем AppController нужно будет обновить.
            method_name = "import_by_mode_key" # <-- Новое имя метода
            method = getattr(self.app_controller, method_name, None)
            if method is None:
                raise AttributeError(f"AppController не имеет метода {method_name}")

            # Вызываем метод с новым ключом
            # --- ИЗМЕНЕНО: Передаем import_mode_key вместо import_type и import_mode ---
            success = method(self.file_path, self.import_mode_key, db_path=self.db_path, progress_callback=internal_progress_callback)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            logger.info(f"Импорт (режим: {self.import_mode_key}) для файла {self.file_path} завершён в потоке {id(QThread.currentThread())}.")

            # Отправляем результат
            self.finished.emit(success, f"Импорт ({self.import_mode_key}) {'успешен' if success else 'неудачен'}.")

        except Exception as e:
            logger.error(f"Ошибка в потоке импорта для файла {self.file_path} (режим: {self.import_mode_key}): {e}", exc_info=True)
            # Отправляем ошибку
            self.finished.emit(False, f"Ошибка импорта: {e}")