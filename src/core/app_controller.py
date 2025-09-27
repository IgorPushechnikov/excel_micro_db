# src/core/app_controller.py
"""
Модуль для центрального контроллера приложения.
Управляет жизненным циклом приложения и координирует работу
между различными менеджерами (DataManager, AnalysisManager и т.д.).
"""

import os
import logging
import sqlite3
from typing import Dict, Any, List, Optional, Tuple, Union
from pathlib import Path

# Импортируем анализатор
# from src.analyzer.logic_documentation import analyze_excel_file # Импорт будет в AnalysisManager

# Импортируем хранилище
from src.storage.base import ProjectDBStorage

# Импортируем экспортёры
# from src.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter as export_with_xlsxwriter # Импорт будет в ExportManager

# Импортируем logger из utils
from src.utils.logger import get_logger

# --- Исключения ---
from src.exceptions.app_exceptions import ProjectError, AnalysisError, ExportError

# Импорты для новых менеджеров (теперь из поддиректории)
# from .controller.data_manager import DataManager # <-- УДАЛЯЕМ ОТСЮДА, чтобы избежать циклического импорта
from .project_manager import ProjectManager # Был перемещен в корень core
# from .controller.format_manager import FormatManager # Пока не реализован
# from .controller.chart_manager import ChartManager # Пока не реализован
# from .controller.analysis_manager import AnalysisManager # Пока не реализован
# from .controller.export_manager import ExportManager # Пока не реализован
# from .controller.node_manager import NodeManager # Пока не реализован

logger = get_logger(__name__)

class AppController:
    """
    Центральный контроллер приложения.

    Координирует работу анализатора, хранилища, процессора и экспортера.
    """

    def __init__(self, project_path: str):
        """
        Инициализирует контроллер для проекта.

        Args:
            project_path (str): Путь к директории проекта.
        """
        self.project_path = project_path
        self.project_db_path = os.path.join(project_path, "project_data.db")
        self.storage: Optional[ProjectDBStorage] = None
        self._current_project_data: Optional[Dict[str, Any]] = None  # Кэш метаданных проекта

        # --- Инициализация менеджеров ---
        # Импортируем DataManager локально, чтобы избежать циклического импорта
        from .controller.data_manager import DataManager # <-- ДОБАВЛЯЕМ СЮДА
        self.data_manager = DataManager(self)

        self.project_manager = ProjectManager(self)
        # self.format_manager = FormatManager(self) # Пока не реализован
        # self.chart_manager = ChartManager(self) # Пока не реализован
        # self.analysis_manager = AnalysisManager(self) # Пока не реализован
        # self.export_manager = ExportManager(self) # Пока не реализован
        # self.node_manager = NodeManager(self) # Пока не реализован

        logger.debug(f"AppController инициализирован для проекта: {project_path}")

    def initialize(self) -> bool:
        """
        Инициализирует контроллер приложения.
        
        Если project_path указан и существует, пытается загрузить проект.
        Если project_path не указан или проект не существует, готов к созданию нового проекта.
        
        Returns:
            bool: True, если инициализация прошла успешно.
        """
        return self.project_manager.initialize()

    @property
    def is_project_loaded(self) -> bool:
        """Проверяет, загружен ли проект."""
        return self.storage is not None

    @property
    def current_project(self) -> Optional[Dict[str, Any]]:
        """Возвращает текущие метаданные проекта."""
        return self._current_project_data

    # --- Управление проектом (делегировано ProjectManager) ---
    def create_project(self, project_path: str) -> bool:
        """Создает новый проект."""
        return self.project_manager.create_project(project_path)

    def create_new_project(self, project_name: str) -> bool:
        """Создает новую структуру проекта."""
        return self.project_manager.create_new_project(project_name)

    def load_project(self) -> bool:
        """Загружает существующий проект."""
        return self.project_manager.load_project()

    def close_project(self):
        """Закрывает текущий проект."""
        self.project_manager.close_project()
        self._current_project_data = None

    def shutdown(self):
        """Полное завершение работы контроллера."""
        self.close_project()
        logger.info("AppController завершил работу.")

    # --- Работа с данными листа (делегировано DataManager) ---
    def get_sheet_data(self, sheet_name: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """Получает данные листа для отображения."""
        return self.data_manager.get_sheet_data(sheet_name)

    def get_sheet_editable_data(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """Получает редактируемые данные листа."""
        return self.data_manager.get_sheet_editable_data(sheet_name)

    def update_sheet_cell_in_project(self, sheet_name: str, row_index: int, column_name: str, new_value: str) -> bool:
        """Обновляет значение ячейки в проекте."""
        return self.data_manager.update_sheet_cell_in_project(sheet_name, row_index, column_name, new_value)

    def update_cell_value(self, sheet_name: str, cell_address: str, new_value: Any) -> bool:
        """Обновляет значение ячейки."""
        return self.data_manager.update_cell_value(sheet_name, cell_address, new_value)

    def get_edit_history(self, sheet_name: Optional[str] = None, limit: Optional[int] = 10) -> List[Dict[str, Any]]:
        """Получает историю редактирования."""
        return self.data_manager.get_edit_history(sheet_name, limit)

    # --- Анализ Excel-файлов (делегировано AnalysisManager - заглушка) ---
    def analyze_excel_file(self, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
        """Анализирует Excel-файл и сохраняет результаты в БД проекта."""
        # Пока вызываем напрямую, но в будущем будет через AnalysisManager
        from src.analyzer.logic_documentation import analyze_excel_file
        if not self.storage:
            logger.error("Проект не загружен. Невозможно выполнить анализ.")
            return False

        if not os.path.exists(file_path):
            logger.error(f"Excel-файл для анализа не найден: {file_path}")
            return False

        logger.info(f"Начало анализа Excel-файла: {file_path}")
        try:
            analysis_results = analyze_excel_file(file_path)
            logger.debug("Анализ Excel-файла завершен.")

            # 2. Сохранение результатов в БД
            # Предполагаем, что analyze_excel_file возвращает данные в формате,
            # который storage может принять
            # Для каждого листа в результатах анализа
            for sheet_data in analysis_results.get("sheets", []):
                sheet_name = sheet_data["name"]
                logger.info(f"Сохранение данных для листа: {sheet_name}")

                # --- Получаем или создаем запись листа в БД ---
                # Это может потребовать отдельного метода в storage
                # Пока делаем это напрямую
                # TODO: Реализовать метод в storage для создания/получения sheet_id

                sheet_id = self._get_or_create_sheet_id(analysis_results.get("project_name", "Unknown"), sheet_name)
                if sheet_id is None:
                    logger.error(f"Не удалось получить/создать ID для листа '{sheet_name}'. Пропущен.")
                    continue

                # --- Сохраняем метаданные листа ---
                metadata_to_save = {
                    "max_row": sheet_data.get("max_row"),
                    "max_column": sheet_data.get("max_column"),
                    "merged_cells": sheet_data.get("merged_cells", [])
                }
                if not self.storage.save_sheet_metadata(sheet_name, metadata_to_save):
                    logger.warning(f"Не удалось сохранить метаданные для листа '{sheet_name}'.")

                # --- Сохраняем объединенные ячейки в отдельную таблицу ---
                merged_cells_list = sheet_data.get("merged_cells", [])
                if merged_cells_list: # Сохраняем только если список не пуст
                    if not self.storage.save_sheet_merged_cells(sheet_id, merged_cells_list):
                        logger.error(f"Не удалось сохранить объединенные ячейки для листа '{sheet_name}' (ID: {sheet_id}).")

                # --- Сохраняем "сырые данные" ---
                if not self.storage.save_sheet_raw_data(sheet_name, sheet_data.get("raw_data", [])):
                    logger.error(f"Не удалось сохранить 'сырые данные' для листа '{sheet_name}'.")

                # --- Сохраняем формулы ---
                if not self.storage.save_sheet_formulas(sheet_id, sheet_data.get("formulas", [])):
                    logger.error(f"Не удалось сохранить формулы для листа '{sheet_name}' (ID: {sheet_id}).")

                # --- Сохраняем стили ---
                if not self.storage.save_sheet_styles(sheet_id, sheet_data.get("styles", [])):
                    logger.error(f"Не удалось сохранить стили для листа '{sheet_name}' (ID: {sheet_id}).")

                # --- Сохраняем диаграммы ---
                if not self.storage.save_sheet_charts(sheet_id, sheet_data.get("charts", [])):
                    logger.error(f"Не удалось сохранить диаграммы для листа '{sheet_name}' (ID: {sheet_id}).")

            logger.info(f"Анализ и сохранение данных из '{file_path}' завершены.")
            return True
        except Exception as e:
            logger.error(f"Ошибка при анализе/сохранении файла '{file_path}': {e}", exc_info=True)
            return False

    def _get_or_create_sheet_id(self, project_name: str, sheet_name: str) -> Optional[int]:
        """
        Получает ID листа из БД или создает новую запись, если лист не существует.

        Это вспомогательный метод, который может быть перенесен в storage в будущем.
        """
        if not self.storage or not self.storage.connection:
            logger.error("Нет подключения к БД для получения/создания sheet_id.")
            return None

        try:
            cursor = self.storage.connection.cursor()

            # Получаем project_id (для MVP предполагаем 1)
            # TODO: Реализовать правильное получение project_id
            project_id = 1

            # Проверяем, существует ли лист
            cursor.execute("SELECT sheet_id FROM sheets WHERE project_id = ? AND name = ?", (project_id, sheet_name))
            result = cursor.fetchone()
            if result:
                logger.debug(f"Лист '{sheet_name}' найден с ID {result[0]}.")
                return result[0]
            else:
                # Создаем новый лист
                logger.debug(f"Лист '{sheet_name}' не найден. Создаем новый.")
                cursor.execute(
                    "INSERT INTO sheets (project_id, name) VALUES (?, ?)",
                    (project_id, sheet_name)
                )
                self.storage.connection.commit()
                new_sheet_id = cursor.lastrowid
                logger.info(f"Создан новый лист '{sheet_name}' с ID {new_sheet_id}.")
                return new_sheet_id
        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при получении/создании sheet_id для '{sheet_name}': {e}")
            return None
        except Exception as e:
            logger.error(f"Неожиданная ошибка при получении/создании sheet_id для '{sheet_name}': {e}", exc_info=True)
            return None

    # --- Экспорт (делегировано ExportManager - заглушка) ---
    def export_results(self, export_type: str, output_path: str) -> bool:
        """Универсальный метод для экспорта результатов проекта."""
        # Пока вызываем напрямую, но в будущем будет через ExportManager
        if export_type.lower() == 'excel':
            return self.export_project(output_path, use_xlsxwriter=True)
        elif export_type.lower() == 'go_excel':
            return self.export_project_with_go(output_path)
        else:
            logger.error(f"Неподдерживаемый тип экспорта: {export_type}")
            return False

    def export_project(self, output_path: str, use_xlsxwriter: bool = True) -> bool:
        """Экспортирует проект в Excel-файл (старый метод)."""
        logger.info(f"Начало экспорта проекта в '{output_path}'. Используется {'xlsxwriter' if use_xlsxwriter else 'openpyxl (отключен)'}."),
        try:
            from src.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter as export_with_xlsxwriter
            success = export_with_xlsxwriter(self.project_db_path, output_path)
            if success:
                logger.info(f"Проект успешно экспортирован в '{output_path}'.")
            else:
                logger.error(f"Ошибка при экспорте проекта в '{output_path}'.")
            return success
        except Exception as e:
            logger.error(f"Неожиданная ошибка при экспорте проекта в '{output_path}': {e}", exc_info=True)
            return False

    def export_project_with_go(self, output_path: str) -> bool:
        """Экспортирует проект в Excel-файл с использованием Go-экспортера."""
        logger.info(f"Начало экспорта проекта в '{output_path}' с использованием Go-экспортера.")
        try:
            # Импортируем GoExporterBridge
            from src.exporter.go_bridge import GoExporterBridge
            
            # Создаем экземпляр моста
            go_bridge = GoExporterBridge(self.storage)

            # --- НОВЫЙ КОД: Подготовка пути для отладочного JSON ---
            output_path_obj = Path(output_path)
            # Определяем путь к папке экспорта (например, та же папка, что и у output файла)
            export_dir = output_path_obj.parent
            # Определяем путь к папке для SQL-дампов и отладочных файлов
            sql_and_debug_dir = export_dir / "sql_export" # Используем ту же папку, что и для SQL
            # Создаем папку, если её нет
            sql_and_debug_dir.mkdir(parents=True, exist_ok=True)
            # Определяем путь к отладочному JSON-файлу
            debug_json_file_path = sql_and_debug_dir / f"{output_path_obj.stem}_debug_data.json"
            logger.debug(f"Debug JSON will be saved to: {debug_json_file_path}")
            # --- КОНЕЦ НОВОГО КОДА ---

            # Выполняем экспорт, передавая путь для отладочного JSON
            # success = go_bridge.export_to_xlsx(Path(output_path)) # Старая строка
            success = go_bridge.export_to_xlsx(Path(output_path), debug_json_path=debug_json_file_path) # Новая строка

            if success:
                logger.info(f"Проект успешно экспортирован в '{output_path}' с помощью Go-экспортера.")
                # --- НОВЫЙ КОД: Сообщение о сохранении отладочного JSON ---
                logger.info(f"Отладочный JSON-файл сохранен: {debug_json_file_path}")
                # --- КОНЕЦ НОВОГО КОДА ---
            else:
                logger.error(f"Ошибка при экспорте проекта в '{output_path}' с помощью Go-экспортера.")
            return success
        except Exception as e:
            logger.error(f"Неожиданная ошибка при экспорте проекта в '{output_path}' с помощью Go-экспортера: {e}", exc_info=True)
            return False


def create_app_controller(project_path: Optional[str] = None) -> AppController:
    """
    Фабричная функция для создания и инициализации экземпляра AppController.

    Args:
        project_path (Optional[str]): Путь к директории проекта. 
                                      Если None, создается контроллер без привязки к проекту.

    Returns:
        AppController: Инициализированный экземпляр контроллера приложения.
    """
    controller = AppController(project_path or "")
    return controller
