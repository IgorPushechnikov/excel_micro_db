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

            # --- ИЗМЕНЕНО: Определение метода AppController через сопоставление ---
            # Определяем метод AppController на основе типа и режима
            # Словарь сопоставления
            method_mapping = {
                ('all', 'openpyxl'): 'import_all_data_from_excel',
                ('raw', 'openpyxl'): 'import_raw_data_from_excel',
                ('styles', 'openpyxl'): 'import_styles_from_excel',
                ('charts', 'openpyxl'): 'import_charts_from_excel',
                ('formulas', 'openpyxl'): 'import_formulas_from_excel',
                ('raw', 'fast_pandas'): 'import_raw_data_fast_with_pandas',
                ('raw', 'selective'): 'import_raw_data_from_excel_selective',
                ('styles', 'selective'): 'import_styles_from_excel_selective',
                ('charts', 'selective'): 'import_charts_from_excel_selective',
                ('formulas', 'selective'): 'import_formulas_from_excel_selective',
                ('raw', 'in_chunks'): 'import_raw_data_from_excel_in_chunks',
                ('all', 'selective'): 'import_all_data_from_excel_selective',
                ('all', 'fast_pandas'): 'import_all_data_from_excel_fast',
                ('all', 'fast'): 'import_all_data_from_excel_fast',
                ('all', 'in_chunks'): 'import_all_data_from_excel_chunks',
                # --- НОВОЕ: Добавлено сопоставление для 'auto' ---
                ('auto', ''): 'import_auto_data_from_excel', # <-- Режим 'auto' не требует дополнительного режима
                # ----------------------------------------------
            }

            # Логика для chunks_openpyxl
            if self.import_type == 'chunks' and self.import_mode == 'openpyxl':
                 # Предполагаем, что chunks_openpyxl означает импорт "сырых данных" частями
                 method_name = 'import_raw_data_from_excel_in_chunks'
            # --- НОВОЕ: Логика для auto ---
            elif self.import_type == 'auto':
                 # Режим 'auto' не требует дополнительного режима
                 method_name = 'import_auto_data_from_excel'
            # --- КОНЕЦ НОВОГО ---
            elif self.import_mode == 'fast_pandas' and self.import_type != 'raw':
                 # Обработка других типов с fast_pandas (если понадобится, сейчас не поддерживается)
                 method_name = method_mapping.get((self.import_type, self.import_mode))
                 if not method_name:
                     raise AttributeError(f"AppController не поддерживает комбинацию import_type={self.import_type} и import_mode={self.import_mode}")
            else:
                 # Используем сопоставление
                 method_name = method_mapping.get((self.import_type, self.import_mode))
                 if not method_name:
                     # Попробуем сопоставление с суффиксом, если не найдено точное
                     if self.import_mode != 'all' and self.import_mode != 'fast':
                         method_name = f"import_{self.import_type}_from_excel_{self.import_mode}"
                     else:
                         method_name = f"import_{self.import_type}_from_excel"

            method = getattr(self.app_controller, method_name, None)
            if method is None:
                raise AttributeError(f"AppController не имеет метода {method_name}")

            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

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
