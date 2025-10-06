# backend/core/controller/export_manager.py

import logging
from typing import Dict, Any, Optional, Callable # <-- Добавлен Callable
from pathlib import Path

# Импортируем функцию экспорта
from backend.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter
from backend.storage.base import ProjectDBStorage # <-- Импортируем ProjectDBStorage
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ExportManager:
    def __init__(self, app_controller):
        """
        Инициализирует ExportManager.
        Теперь не хранит ссылку на storage, так как будет создавать его в потоке.

        Args:
            app_controller: Экземпляр AppController (может понадобиться для других целей, но не для storage).
        """
        self.app_controller = app_controller
        logger.debug("ExportManager инициализирован.")

    def perform_export(self, export_type: str, output_path: str, db_path: str, options: Optional[Dict[str, Any]] = None, progress_callback: Optional[Callable[[int, str], None]] = None) -> bool: # <-- ИЗМЕНЕНО: Добавлен db_path и progress_callback
        """
        Выполняет экспорт данных проекта в файл.
        В текущей реализации поддерживает только 'excel' через xlsxwriter.
        Создаёт собственное соединение с БД в текущем потоке.

        Args:
            export_type (str): Тип экспорта (например, 'excel').
            output_path (str): Путь к выходному файлу.
            db_path (str): Путь к файлу БД проекта (.db).
            options (Optional[Dict[str, Any]]): Опции экспорта.
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.

        Returns:
            bool: True, если экспорт прошёл успешно, иначе False.
        """
        # Проверяем тип экспорта
        if export_type.lower() != 'excel':
            logger.error(f"ExportManager: Тип экспорта '{export_type}' не поддерживается.")
            return False

        # Создаём ProjectDBStorage с указанным db_path ВНУТРИ текущего потока
        storage = ProjectDBStorage(db_path)
        if not storage.connect():
            logger.error(f"ExportManager: Не удалось подключиться к БД проекта {db_path} в потоке {id(__import__('threading').current_thread())}.")
            return False

        try:
            logger.info(f"ExportManager: Начало экспорта в '{output_path}' (тип: {export_type}, БД: {db_path})")
            # Проверяем, что файл БД существует
            if not Path(db_path).exists():
                logger.error(f"ExportManager: Файл БД проекта не найден: {db_path}")
                return False

            # --- НОВОЕ: Обновление прогресса в начале ---
            if progress_callback:
                progress_callback(0, f"Экспорт в {output_path}...")
            # --- КОНЕЦ НОВОГО ---

            # Вызов функции экспорта из xlsxwriter_exporter
            # Передаём путь к БД и путь к выходному файлу
            # И передаём progress_callback
            success = export_project_xlsxwriter(db_path, output_path, progress_callback=progress_callback) # <-- ИЗМЕНЕНО: Добавлен progress_callback

            # --- НОВОЕ: Обновление прогресса в конце или ошибке ---
            if progress_callback:
                if success:
                    progress_callback(100, "Экспорт завершён.")
                else:
                    progress_callback(0, "Экспорт не удался.")
            # --- КОНЕЦ НОВОГО ---

            if success:
                logger.info(f"ExportManager: Экспорт в '{output_path}' завершён успешно.")
            else:
                logger.error(f"ExportManager: Ошибка при экспорте в '{output_path}'.")
            
            return success

        except Exception as e:
            logger.error(f"ExportManager: Ошибка при экспорте в '{output_path}': {e}", exc_info=True)
            # --- НОВОЕ: Обновление прогресса при ошибке ---
            if progress_callback:
                progress_callback(0, f"Ошибка экспорта: {e}")
            # --- КОНЕЦ НОВОГО ---
            return False
        finally:
            # 3. Закрытие соединения с БД в текущем потоке
            storage.disconnect()
            logger.debug(f"ExportManager: Соединение с БД {db_path} закрыто в потоке {id(__import__('threading').current_thread())}.")
