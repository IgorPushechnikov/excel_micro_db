# backend/core/app_controller.py

import os
import logging
import sqlite3
from typing import Dict, Any, List, Optional, Tuple, Union, Callable # <-- ДОБАВЛЕНО: Callable
from pathlib import Path # <-- ДОБАВЛЕНО: Импорт Path из pathlib

# Импортируем анализатор
# from analyzer.logic_documentation import analyze_excel_file # Импорт будет в AnalysisManager

# Импортируем хранилище
from backend.storage.base import ProjectDBStorage # <-- ИСПРАВЛЕНО: Импорт теперь из backend.storage

# Импортируем экспортёры
# from exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter as export_with_xlsxwriter # Импорт будет в ExportManager

# Импортируем logger из utils
# ИСПРАВЛЕНО: Корректный путь к logger внутри backend
from backend.utils.logger import get_logger, set_logging_enabled, is_logging_enabled # <-- ИСПРАВЛЕНО: было from utils.logger

# --- Исключения ---
from backend.exceptions.app_exceptions import ProjectError, AnalysisError, ExportError

# Импорты для новых менеджеров (теперь из поддиректории)
# from .controller.data_manager import DataManager # <-- УДАЛЯЕМ ОТСЮДА, чтобы избежать циклического импорта
from .project_manager import ProjectManager # Был перемещен в корень core
# from .controller.format_manager import FormatManager # Пока не реализован
# from .controller.chart_manager import ChartManager # Пока не реализован
from .controller.analysis_manager import AnalysisManager # <-- НОВОЕ: Импорт AnalysisManager
from .controller.export_manager import ExportManager # <-- НОВОЕ: Импорт ExportManager
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
        # --- НОВОЕ: Атрибут для обработчика логов проекта ---
        self._project_log_handler: Optional[logging.FileHandler] = None
        # ================================================
        # --- НОВОЕ: Атрибут для хранения пути к последнему импортированному файлу ---
        self.last_imported_file_path: Optional[str] = None
        # ================================================

        # --- Инициализация менеджеров ---
        # Импортируем DataManager локально, чтобы избежать циклического импорта
        from .controller.data_manager import DataManager # <-- ДОБАВЛЯЕМ СЮДА
        self.data_manager = DataManager(self)

        self.project_manager = ProjectManager(self)
        # self.format_manager = FormatManager(self) # Пока не реализован
        # self.chart_manager = ChartManager(self) # Пока не реализован
        self.analysis_manager = AnalysisManager(self) # <-- НОВОЕ: Инициализация AnalysisManager
        self.export_manager = ExportManager(self) # <-- НОВОЕ: Инициализация ExportManager
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
        success = self.project_manager.create_project(project_path)
        if success:
            # --- НОВОЕ: Настройка логирования проекта ---
            self._setup_project_logging(project_path)
            # ==========================================
            # --- ИСПРАВЛЕНО: Обновление self.project_db_path при создании проекта ---
            self.project_path = project_path
            self.project_db_path = os.path.join(project_path, "project_data.db")
            # ----------------------------------------------------------------------
        return success

    def create_new_project(self, project_name: str) -> bool:
        """Создает новую структуру проекта."""
        return self.project_manager.create_new_project(project_name)

    def load_project(self, project_path: Optional[str] = None) -> bool:
        """Загружает существующий проект."""
        # Используем переданный путь или сохраненный
        load_path = project_path or self.project_path
        success = self.project_manager.load_project(load_path)
        if success:
            # --- НОВОЕ: Настройка логирования проекта ---
            self._setup_project_logging(load_path)
            # ==========================================
            # --- ИСПРАВЛЕНО: Обновление self.project_db_path при загрузке проекта ---
            self.project_path = load_path
            self.project_db_path = os.path.join(load_path, "project_data.db")
            # ----------------------------------------------------------------------
        return success

    def close_project(self):
        """Закрывает текущий проект."""
        # --- НОВОЕ: Удаление обработчика логов проекта ---
        self._remove_project_logging()
        # ==============================================
        self.project_manager.close_project()
        # --- ИСПРАВЛЕНО: Обнуление self.project_db_path при закрытии проекта ---
        self.project_db_path = ""
        # ----------------------------------------------------------------------
        self._current_project_data = None
        # --- ОБНУЛЕНИЕ last_imported_file_path при закрытии проекта ---
        self.last_imported_file_path = None
        # -------------------------------------------------------------

    def shutdown(self):
        """Полное завершение работы контроллера."""
        self.close_project()
        logger.info("AppController завершил работу.")

    # --- НОВОЕ: Метод для анализа Excel файла ---
    def analyze_excel_file(self, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
        """
        Анализирует Excel-файл и сохраняет результаты в БД проекта через AnalysisManager.

        Args:
            file_path (str): Путь к Excel-файлу для анализа.
            options (Optional[Dict[str, Any]]): Опции анализа.

        Returns:
            bool: True, если анализ успешен, иначе False.
        """
        if not self.storage:
            logger.error("Проект не загружен. Невозможно выполнить анализ.")
            return False

        logger.info(f"AppController: Запуск анализа файла {file_path} через AnalysisManager.")
        # Делегирование AnalysisManager
        return self.analysis_manager.perform_analysis(file_path, options)
    # --- КОНЕЦ НОВОГО ---

    # --- НОВОЕ: Метод для экспорта проекта ---
    def export_results(self, export_type: str, output_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
        """
        Экспортирует результаты проекта через ExportManager.

        Args:
            export_type (str): Тип экспорта (например, 'excel').
            output_path (str): Путь к выходному файлу.
            options (Optional[Dict[str, Any]]): Опции экспорта.

        Returns:
            bool: True, если экспорт успешен, иначе False.
        """
        if not self.storage:
            logger.error("Проект не загружен. Невозможно выполнить экспорт.")
            return False

        logger.info(f"AppController: Запуск экспорта в {output_path} (тип: {export_type}) через ExportManager.")
        # Делегирование ExportManager
        return self.export_manager.perform_export(export_type, output_path, options)
    # --- КОНЕЦ НОВОГО ---

    # --- Работа с данными листа (делегировано DataManager) ---
    def get_sheet_data(self, sheet_name: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """Получает данные листа для отображения."""
        return self.data_manager.get_sheet_data(sheet_name)

    def get_sheet_editable_data(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """Получает редактируемые данные листа."""
        return self.data_manager.get_sheet_editable_data(sheet_name)

    def get_sheet_raw_data(self, sheet_name: str) -> Optional[List[Dict[str, Any]]]:
        """Получает "сырые" данные листа (включая формулы, стили и т.д.)."""
        return self.data_manager.get_sheet_raw_data(sheet_name)

    def update_sheet_cell_in_project(self, sheet_name: str, row_index: int, column_name: str, new_value: str) -> bool:
        """Обновляет значение ячейки в проекте."""
        return self.data_manager.update_sheet_cell_in_project(sheet_name, row_index, column_name, new_value)

    def update_cell_value(self, sheet_name: str, cell_address: str, new_value: Any) -> bool:
        """Обновляет значение ячейки."""
        return self.data_manager.update_cell_value(sheet_name, cell_address, new_value)

    def get_edit_history(self, sheet_name: Optional[str] = None, limit: Optional[int] = 10) -> List[Dict[str, Any]]:
        """Получает историю редактирования."""
        return self.data_manager.get_edit_history(sheet_name, limit)

    def get_sheet_names(self) -> List[str]:
        """
        Получает список имен листов из текущего проекта.

        Returns:
            List[str]: Список имен листов. Возвращает пустой список в случае ошибки или отсутствия подключения.
        """
        if not self.storage or not self.storage.connection:
            logger.error("Нет подключения к БД для получения списка листов.")
            return []

        try:
            cursor = self.storage.connection.cursor()
            # Используем правильное имя столбца 'name' из таблицы 'sheets'
            cursor.execute("SELECT name FROM sheets ORDER BY name;")
            rows = cursor.fetchall()
            sheet_names = [row[0] for row in rows]
            logger.info(f"Получено {len(sheet_names)} имен листов из БД.")
            return sheet_names
        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при получении списка листов: {e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка при получении списка листов: {e}", exc_info=True)
            return []

    # --- НОВОЕ: Метод для переименования листа ---
    def rename_sheet(self, old_name: str, new_name: str) -> bool:
        """
        Переименовывает лист в проекте.

        Args:
            old_name (str): Текущее имя листа.
            new_name (str): Новое имя листа.

        Returns:
            bool: True, если переименование успешно, иначе False.
        """
        if not self.storage:
            logger.error("Проект не загружен. Невозможно переименовать лист.")
            return False

        if not old_name or not new_name:
            logger.error("Имена листов (старое и новое) не могут быть пустыми.")
            return False

        # Получаем project_id. В MVP предполагаем, что это 1, но лучше бы получить из метаданных проекта.
        # Для простоты используем 1, как в других местах.
        project_id = 1

        logger.info(f"Попытка переименования листа '{old_name}' в '{new_name}' для проекта ID {project_id}.")
        success = self.storage.rename_sheet(project_id, old_name, new_name)
        if success:
            logger.info(f"Лист '{old_name}' успешно переименован в '{new_name}'.")
            # Возможно, нужно обновить внутренние кэши или уведомить другие компоненты
            # о смене имени, но для базовой реализации этого достаточно.
        else:
            logger.error(f"Не удалось переименовать лист '{old_name}' в '{new_name}'.")
        return success
    # --- КОНЕЦ НОВОГО ---

    # --- НОВОЕ: Методы для управления глобальным логированием ---
    def set_logging_enabled(self, enabled: bool):
        """
        Включает или отключает логирование для всего приложения.

        Args:
            enabled (bool): True для включения, False для отключения.
        """
        set_logging_enabled(enabled)

    def is_logging_enabled(self) -> bool:
        """
        Проверяет, включено ли логирование.

        Returns:
            bool: True, если логирование включено, иначе False.
        """
        return is_logging_enabled()
    # --- КОНЕЦ НОВОГО ---

    # --- ВНУТРЕННИЕ МЕТОДЫ ДЛЯ УПРАВЛЕНИЯ ЛОГИРОВАНИЕМ ПРОЕКТА ---
    def _setup_project_logging(self, project_path: str):
        """
        Настраивает FileHandler для логирования в файл проекта.
        """
        # Путь к файлу лога проекта
        log_file_path = os.path.join(project_path, "logs", f"app_controller_{os.path.basename(project_path)}.log")
        os.makedirs(os.path.dirname(log_file_path), exist_ok=True)

        # Создаём FileHandler
        self._project_log_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
        # Создаём форматтер
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        self._project_log_handler.setFormatter(formatter)
        # Добавляем FileHandler к логгеру AppController
        logger.addHandler(self._project_log_handler)
        logger.info(f"Логирование проекта настроено: {log_file_path}")

    def _remove_project_logging(self):
        """
        Удаляет FileHandler логирования проекта.
        """
        if self._project_log_handler:
            logger.info(f"Удаление обработчика логов проекта: {self._project_log_handler.baseFilename}")
            logger.removeHandler(self._project_log_handler)
            self._project_log_handler.close()
            self._project_log_handler = None
        else:
            logger.debug("Обработчик логов проекта не был установлен.")
    # --- КОНЕЦ ВНУТРЕННИХ МЕТОДОВ ЛОГИРОВАНИЯ ---

    # --- Существующие методы импорта (заглушки или делегирование) ---

    # --- Существующие методы экспорта (заглушки или делегирование) ---

    # --- Внутренние методы для управления менеджерами ---

    # --- Вспомогательные методы ---


# --- Фабричная функция ---
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
