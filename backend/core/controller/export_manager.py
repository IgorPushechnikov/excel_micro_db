# backend/core/controller/export_manager.py

import logging
from typing import Dict, Any, Optional
from pathlib import Path

# Импортируем функцию экспорта
from backend.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter
from backend.storage.base import ProjectDBStorage
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ExportManager:
    def __init__(self, app_controller):
        """
        Инициализирует ExportManager.

        Args:
            app_controller: Экземпляр AppController, откуда будет получен storage.
        """
        self.app_controller = app_controller
        logger.debug("ExportManager инициализирован.")

    def perform_export(self, export_type: str, output_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
        """
        Выполняет экспорт данных проекта в файл.
        В текущей реализации поддерживает только 'excel' через xlsxwriter.

        Args:
            export_type (str): Тип экспорта (например, 'excel').
            output_path (str): Путь к выходному файлу.
            options (Optional[Dict[str, Any]]): Опции экспорта.

        Returns:
            bool: True, если экспорт прошёл успешно, иначе False.
        """
        storage = self.app_controller.storage
        if not storage:
            logger.error("Storage не загружен в AppController. Невозможно выполнить экспорт.")
            return False

        # Проверяем тип экспорта
        if export_type.lower() != 'excel':
            logger.error(f"Тип экспорта '{export_type}' не поддерживается.")
            return False

        try:
            logger.info(f"Начало экспорта в '{output_path}' (тип: {export_type})")
            project_db_path = self.app_controller.project_db_path
            # Проверяем, что файл БД существует
            if not Path(project_db_path).exists():
                logger.error(f"Файл БД проекта не найден: {project_db_path}")
                return False

            # Вызов функции экспорта из xlsxwriter_exporter
            # Передаём storage.connection (путь к БД) и путь к выходному файлу
            success = export_project_xlsxwriter(project_db_path, output_path)

            if success:
                logger.info(f"Экспорт в '{output_path}' завершён успешно.")
            else:
                logger.error(f"Ошибка при экспорте в '{output_path}'.")
            
            return success

        except Exception as e:
            logger.error(f"Ошибка при экспорте в '{output_path}': {e}", exc_info=True)
            return False
