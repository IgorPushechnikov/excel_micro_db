# backend/core/controller/analysis_manager.py

import logging
from typing import Dict, Any, Optional
from backend.analyzer.logic_documentation import analyze_excel_file as run_analysis
from backend.storage.base import ProjectDBStorage
from backend.utils.logger import get_logger

logger = get_logger(__name__)

class AnalysisManager:
    def __init__(self, app_controller):
        """
        Инициализирует AnalysisManager.

        Args:
            app_controller: Экземпляр AppController, откуда будет получен storage.
        """
        self.app_controller = app_controller
        logger.debug("AnalysisManager инициализирован.")

    def perform_analysis(self, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
        """
        Выполняет анализ Excel-файла и сохраняет результаты в БД проекта.

        Args:
            file_path (str): Путь к Excel-файлу для анализа.
            options (Optional[Dict[str, Any]]): Опции анализа (например, max_rows, include_formulas).

        Returns:
            bool: True, если анализ и сохранение прошли успешно, иначе False.
        """
        storage = self.app_controller.storage
        if not storage:
            logger.error("Storage не загружен в AppController. Невозможно выполнить анализ.")
            return False

        try:
            logger.info(f"Начало анализа Excel-файла: {file_path}")

            # 1. Вызов анализатора
            analysis_results = run_analysis(file_path)

            # 2. Обработка и сохранение результатов
            for sheet_info in analysis_results.get("sheets", []):
                sheet_name = sheet_info["name"]
                raw_data = sheet_info["raw_data"]
                formulas = sheet_info["formulas"]
                styles = sheet_info["styles"]
                charts = sheet_info["charts"]
                merged_cells = sheet_info["merged_cells"]
                max_row = sheet_info.get("max_row")
                max_col = sheet_info.get("max_column")

                logger.debug(f"Обработка листа: {sheet_name}")

                # Сохранение информации о листе
                sheet_id = storage.save_sheet(project_id=1, sheet_name=sheet_name, max_row=max_row, max_column=max_col)
                if sheet_id is None:
                    logger.error(f"Не удалось получить sheet_id для {sheet_name}. Прерывание анализа.")
                    return False

                # Сохранение данных
                if not storage.save_sheet_raw_data(sheet_name, raw_data):
                    logger.error(f"Ошибка сохранения raw_data для {sheet_name}")
                    return False
                if not storage.save_sheet_formulas(sheet_id, formulas):
                    logger.error(f"Ошибка сохранения формул для {sheet_name}")
                    return False
                if not storage.save_sheet_styles(sheet_id, styles):
                    logger.error(f"Ошибка сохранения стилей для {sheet_name}")
                    return False
                if not storage.save_sheet_charts(sheet_id, charts):
                    logger.error(f"Ошибка сохранения диаграмм для {sheet_name}")
                    return False
                if not storage.save_sheet_merged_cells(sheet_id, merged_cells):
                    logger.error(f"Ошибка сохранения объединенных ячеек для {sheet_name}")
                    return False

            logger.info(f"Анализ файла {file_path} завершен успешно.")
            return True

        except Exception as e:
            logger.error(f"Ошибка при анализе файла {file_path}: {e}", exc_info=True)
            return False
