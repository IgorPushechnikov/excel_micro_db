# backend/core/app_controller.py

import os
import logging
import sqlite3
from typing import Dict, Any, List, Optional, Tuple, Union
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

    # --- НОВОЕ: Метод для обновления last_imported_file_path ---
    def _update_last_imported_file_path(self, file_path: str):
        """
        Обновляет атрибут last_imported_file_path.

        Args:
            file_path (str): Путь к импортированному файлу.
        """
        self.last_imported_file_path = file_path
        logger.debug(f"Путь к последнему импортированному файлу обновлён: {file_path}")

    # --- Анализ Excel-файлов (делегировано AnalysisManager - заглушка) ---
    def analyze_excel_file(self, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
        """Анализирует Excel-файл и сохраняет результаты в БД проекта."""
        # Пока вызываем напрямую, но в будущем будет через AnalysisManager
        from backend.analyzer.logic_documentation import analyze_excel_file # <-- ИСПРАВЛЕНО: Импорт теперь из backend.analyzer
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
                sheet_name = sheet_data["name"] # <-- ВОЗВРАЩЕНО: присваивание sheet_name внутри цикла
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

            # --- НОВОЕ: Обновляем last_imported_file_path после успешного анализа ---
            self._update_last_imported_file_path(file_path)
            # -------------------------------------------------------------------------

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
            # Новый питоновский экспорт через xlsxwriter
            return self.export_project_with_xlsxwriter(output_path)
        else:
            logger.error(f"Неподдерживаемый тип экспорта: {export_type}")
            return False

    def export_project(self, output_path: str, use_xlsxwriter: bool = True) -> bool:
        """Экспортирует проект в Excel-файл (старый метод)."""
        logger.info(f"Начало экспорта проекта в '{output_path}'. Используется {'xlsxwriter' if use_xlsxwriter else 'openpyxl (отключен)'}.")
        try:
            from backend.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter as export_with_xlsxwriter # <-- ИСПРАВЛЕНО: Импорт теперь из backend.exporter
            success = export_with_xlsxwriter(self.project_db_path, output_path)
            if success:
                logger.info(f"Проект успешно экспортирован в '{output_path}'.")
            else:
                logger.error(f"Ошибка при экспорте проекта в '{output_path}'.")
            return success
        except Exception as e:
            logger.error(f"Неожиданная ошибка при экспорте проекта в '{output_path}': {e}", exc_info=True)
            return False

    def export_project_with_xlsxwriter(self, output_path: str) -> bool:
        """Экспортирует проект в Excel-файл с использованием Python-экспортера (xlsxwriter)."""
        # --- НОВОЕ: Проверка и добавление расширения .xlsx ---
        output_path_obj = Path(output_path)
        if output_path_obj.suffix.lower() != '.xlsx':
            output_path_obj = output_path_obj.with_suffix('.xlsx')
            output_path = str(output_path_obj)
            logger.info(f"К пути экспорта добавлено расширение '.xlsx'. Новый путь: {output_path}")
        # ----------------------------------------------------

        logger.info(f"Начало экспорта проекта в '{output_path}' с использованием Python-экспортера (xlsxwriter).")
        try:
            # Импортируем xlsxwriter_exporter
            from backend.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter # <-- ИСПРАВЛЕНО: Импорт теперь из backend.exporter
            
            # Выполняем экспорт
            success = export_project_xlsxwriter(self.project_db_path, output_path)

            if success:
                logger.info(f"Проект успешно экспортирован в '{output_path}' с помощью Python-экспортера (xlsxwriter).")
            else:
                logger.error(f"Ошибка при экспорте проекта в '{output_path}' с помощью Python-экспортера (xlsxwriter).")
            return success
        except Exception as e:
            logger.error(f"Неожиданная ошибка при экспорте проекта в '{output_path}' с помощью Python-экспортера (xlsxwriter): {e}", exc_info=True)
            return False

    # --- НОВОЕ: Методы для настройки и удаления логирования проекта ---
    def _setup_project_logging(self, project_path: str):
        """
        Настраивает дублирующий лог-файл в папке проекта.
        """
        try:
            log_dir = Path(project_path) / "logs"
            log_dir.mkdir(exist_ok=True)  # Создаем папку logs, если её нет
            log_file = log_dir / "project_gui.log"

            # Создаем обработчик
            self._project_log_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
            # Создаем форматтер, скопированный из utils.logger
            # LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            self._project_log_handler.setFormatter(formatter)

            # Добавляем обработчик к корневому логгеру или к логгеру приложения
            # Выберем корневой, чтобы захватить все сообщения
            root_logger = logging.getLogger()
            root_logger.addHandler(self._project_log_handler)

            logger.info(f"Логирование в файл проекта '{log_file}' настроено.")
        except Exception as e:
            logger.error(f"Ошибка при настройке логирования в файл проекта: {e}", exc_info=True)

    def _remove_project_logging(self):
        """
        Удаляет обработчик логов проекта при закрытии проекта.
        """
        if self._project_log_handler:
            try:
                root_logger = logging.getLogger()
                root_logger.removeHandler(self._project_log_handler)
                self._project_log_handler.close()  # Закрываем файловый дескриптор
                logger.info(f"Логирование в файл проекта '{self._project_log_handler.baseFilename}' удалено.")
            except Exception as e:
                logger.error(f"Ошибка при удалении обработчика логов проекта: {e}", exc_info=True)
            finally:
                self._project_log_handler = None
        else:
            logger.debug("Обработчик логов проекта не был установлен.")
    # ================================================================

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

    # --- НОВОЕ: Методы для импорта по типам и режимам (обновлены) ---

    def import_raw_data_from_excel(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только "сырые" данные (значения ячеек) из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_raw_data_from_excel as import_func
        return import_func(storage_to_use, file_path, options)

    def import_styles_from_excel(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только стили из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_styles_from_excel as import_func
        return import_func(storage_to_use, file_path, options)

    def import_charts_from_excel(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только диаграммы из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_charts_from_excel as import_func
        return import_func(storage_to_use, file_path, options)

    def import_formulas_from_excel(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только формулы из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_formulas_from_excel as import_func
        return import_func(storage_to_use, file_path, options)

    def import_raw_data_from_excel_selective(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только "сырые" данные выборочно из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_raw_data_from_excel_selective as import_func
        return import_func(storage_to_use, file_path, options)

    def import_styles_from_excel_selective(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только стили выборочно из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_styles_from_excel_selective as import_func
        return import_func(storage_to_use, file_path, options)

    def import_charts_from_excel_selective(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только диаграммы выборочно из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_charts_from_excel_selective as import_func
        return import_func(storage_to_use, file_path, options)

    def import_formulas_from_excel_selective(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует только формулы выборочно из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_formulas_from_excel_selective as import_func
        return import_func(storage_to_use, file_path, options)

    def import_raw_data_from_excel_in_chunks(self, file_path: str, chunk_options: Dict[str, Any], db_path: Optional[str] = None) -> bool:
        """
        Импортирует только "сырые" данные частями из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            chunk_options (Dict[str, Any]): Опции для разбиения на части.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_raw_data_from_excel_in_chunks as import_func
        return import_func(storage_to_use, file_path, chunk_options)

    def import_styles_from_excel_in_chunks(self, file_path: str, chunk_options: Dict[str, Any], db_path: Optional[str] = None) -> bool:
        """
        Импортирует только стили частями из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            chunk_options (Dict[str, Any]): Опции для разбиения на части.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_styles_from_excel_in_chunks as import_func
        return import_func(storage_to_use, file_path, chunk_options)

    def import_charts_from_excel_in_chunks(self, file_path: str, chunk_options: Dict[str, Any], db_path: Optional[str] = None) -> bool:
        """
        Импортирует только диаграммы частями из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            chunk_options (Dict[str, Any]): Опции для разбиения на части.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_charts_from_excel_in_chunks as import_func
        return import_func(storage_to_use, file_path, chunk_options)

    def import_formulas_from_excel_in_chunks(self, file_path: str, chunk_options: Dict[str, Any], db_path: Optional[str] = None) -> bool:
        """
        Импортирует только формулы частями из Excel-файла.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            chunk_options (Dict[str, Any]): Опции для разбиения на части.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_formulas_from_excel_in_chunks as import_func
        return import_func(storage_to_use, file_path, chunk_options)

    def import_raw_data_fast_with_pandas(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Быстро импортирует только "сырые" данные (значения ячеек) из Excel-файла с помощью pandas.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: True, если импорт успешен, иначе False.
        """
        storage_to_use = ProjectDBStorage(db_path) if db_path else self.storage
        if not storage_to_use:
            logger.error("Экземпляр ProjectDBStorage не предоставлен и не загружен проект. Невозможно выполнить импорт.")
            return False

        from .app_controller_data_import import import_raw_data_fast_with_pandas as import_func
        return import_func(storage_to_use, file_path, options)

    # --- НОВОЕ: Заглушки для недостающих методов ---
    def import_all_data_from_excel(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует все типы данных (сырые, стили, диаграммы, формулы) из Excel-файла.
        Пока не реализовано.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: False, так как функция не реализована.
        """
        logger.warning("Метод 'import_all_data_from_excel' не реализован.")
        return False

    def import_raw_data_pandas_from_excel(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Быстро импортирует только "сырые" данные (значения ячеек) из Excel-файла с помощью pandas.
        Это дублирует 'import_raw_data_fast_with_pandas'.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: Результат вызова 'import_raw_data_fast_with_pandas'.
        """
        logger.info("Вызов 'import_raw_data_pandas_from_excel' перенаправлен на 'import_raw_data_fast_with_pandas'.")
        return self.import_raw_data_fast_with_pandas(file_path, options, db_path)

    def import_all_data_from_excel_selective(self, file_path: str, options: Optional[Dict[str, Any]] = None, db_path: Optional[str] = None) -> bool:
        """
        Импортирует все типы данных выборочно из Excel-файла.
        Пока не реализовано.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            options (Optional[Dict[str, Any]]): Опции импорта.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: False, так как функция не реализована.
        """
        logger.warning("Метод 'import_all_data_from_excel_selective' не реализован.")
        return False

    def import_all_data_from_excel_chunks(self, file_path: str, chunk_options: Dict[str, Any], db_path: Optional[str] = None) -> bool:
        """
        Импортирует все типы данных частями из Excel-файла.
        Пока не реализовано.

        Args:
            file_path (str): Путь к Excel-файлу для импорта.
            chunk_options (Dict[str, Any]): Опции для разбиения на части.
            db_path (Optional[str]): Путь к БД проекта. Если None, использует self.storage.

        Returns:
            bool: False, так как функция не реализована.
        """
        logger.warning("Метод 'import_all_data_from_excel_chunks' не реализован.")
        return False

    # --- КОНЕЦ НОВОГО ---


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
