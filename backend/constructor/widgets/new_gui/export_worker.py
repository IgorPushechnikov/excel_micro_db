# backend/constructor/widgets/new_gui/export_worker.py
"""
Вспомогательный класс для выполнения экспорта в отдельном потоке.
"""

import logging
from pathlib import Path
from PySide6.QtCore import QThread, Signal

from backend.core.app_controller import AppController
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ExportWorker(QThread):
    """
    Рабочий поток для выполнения экспорта данных через AppController.
    """
    finished = Signal(bool, str)  # (успех/ошибка, сообщение)
    progress = Signal(int, str)   # (значение, сообщение) - если AppController будет передавать прогресс

    def __init__(self, app_controller: AppController, output_path: str, export_type: str = 'excel'):
        super().__init__()
        self.app_controller = app_controller
        self.output_path = output_path
        self.export_type = export_type
        # --- НОВОЕ: Получаем путь к БД из app_controller ---
        self.db_path = app_controller.project_db_path
        # --- КОНЕЦ НОВОГО ---

    def run(self):
        """
        Запускает экспорт в отдельном потоке.
        """
        try:
            logger.info(f"Начало экспорта (тип: {self.export_type}) в файл {self.output_path} в потоке {id(QThread.currentThread())}")

            # --- НОВОЕ: Создаём функцию-обёртку для прогресса ---
            def progress_callback(value, message):
                # Проверяем, не был ли отменён диалог прогресса
                # (это нужно делать внутри run, так как сигнал приходит асинхронно)
                # QProgressDialog не предоставляет простого способа проверки отмены из потока
                # Можно передать флаг отмены, но это усложнит код.
                # Пока просто эмитим сигнал.
                self.progress.emit(value, message)
            # --- КОНЕЦ НОВОГО ---

            # Вызываем экспорт через контроллер, передаём db_path и callback
            # success = self.app_controller.export_results(export_type=self.export_type, output_path=self.output_path, db_path=self.db_path)
            # Теперь вызываем с progress_callback
            success = self.app_controller.export_results(
                export_type=self.export_type,
                output_path=self.output_path,
                db_path=self.db_path,
                progress_callback=progress_callback # <-- Передаём callback
            )

            logger.info(f"Экспорт (тип: {self.export_type}) в файл {self.output_path} завершён в потоке {id(QThread.currentThread())}.")

            # Отправляем результат
            self.finished.emit(success, f"Экспорт ({self.export_type}) {'успешен' if success else 'неудачен'}.")

        except Exception as e:
            logger.error(f"Ошибка в потоке экспорта в файл {self.output_path} (тип: {self.export_type}): {e}", exc_info=True)
            # Отправляем ошибку
            self.finished.emit(False, f"Ошибка экспорта: {e}")

