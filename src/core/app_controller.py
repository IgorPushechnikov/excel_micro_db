# src/core/app_controller.py

import os
from pathlib import Path
import sqlite3
from typing import Dict, Any, List, Optional, Tuple

# --- ИНТЕГРАЦИЯ: Импорты компонентов ---
from src.storage.base import ProjectDBStorage
from src.analyzer.logic_documentation import analyze_workbook_logic
from src.exporter.direct_db_exporter import export_project_db_to_excel
from src.core.project_manager import ProjectManager
from src.utils.logger import setup_logger
from src.exceptions.app_exceptions import (
    AppException,
    ProjectNotInitializedError,
    ProjectLoadError,
    AnalysisError,
    ExportError,
    ProcessingError,
)

# Получаем логгер для этого модуля
logger = setup_logger(__name__)

class AppController:
    """
    Центральный контроллер приложения Excel Micro DB.

    Отвечает за:
    - Инициализацию и управление компонентами приложения
    - Координацию между модулями (analyzer, processor, storage, exporter, constructor)
    - Управление состоянием приложения
    - Обработку ошибок на уровне приложения
    """

    def __init__(self, project_path: Optional[str] = None, storage: Optional[ProjectDBStorage] = None, log_level: str = "INFO"):
        """
        Инициализация контроллера приложения.

        Args:
            project_path (Optional[str]): Путь к директории проекта
            storage (Optional[ProjectDBStorage]): Экземпляр ProjectDBStorage
            log_level (str): Уровень логирования.
        """
        logger.debug(f"Инициализация AppController для проекта: {project_path}")
        self.project_path = Path(project_path) if project_path else None
        self.storage = storage
        self.log_level = log_level
        self.logger = setup_logger(self.__class__.__name__, log_level=log_level)
        self.current_project_info = None # <-- Информация о метаданных проекта
        self.is_initialized = False # <-- Статус инициализации AppController
        logger.debug("AppController инициализирован")

    def initialize(self):
        """Инициализация внутреннего состояния AppController."""
        try:
            self.logger.info("Инициализация внутреннего состояния AppController...")
            if not self.storage:
                 self.logger.error("Storage не предоставлен для инициализации AppController.")
                 return False
            # Инициализируем схему БД, если нужно
            self.storage.initialize_schema()
            # Пытаемся загрузить метаданные проекта
            self.current_project_info = self.storage.load_project_metadata() # <-- Загружаем метаданные
            if self.current_project_info:
                self.logger.info(f"Метаданные проекта загружены: {self.current_project_info.get('name', 'Unknown')}")
            else:
                self.logger.info("Метаданные проекта не найдены, будет создан новый проект.")

            self.is_initialized = True
            self.logger.info("AppController инициализирован успешно.")
            return True
        except Exception as e:
            self.logger.error(f"Ошибка инициализации AppController: {e}", exc_info=True)
            self.is_initialized = False
            return False

    # --- Управление проектом (через ProjectManager или Storage) ---

    def init_project(self):
        """Инициализация (создание) проекта через ProjectManager."""
        # Используем ProjectManager для создания структуры проекта на диске/в БД
        # Предполагается, что ProjectManager знает, как создать файлы/папки и записать начальные метаданные в БД
        # через self.storage.
        self.logger.info("Инициализация проекта через ProjectManager...")
        try:
            # ProjectManager теперь получает storage, а не управляет им сам
            project_manager = ProjectManager(log_level=self.log_level)
            # Вызываем метод ProjectManager, передавая ему путь и ссылку на storage
            success = project_manager.create_project_structure(str(self.project_path), self.storage)
            if success:
                self.logger.info("Проект успешно инициализирован через ProjectManager.")
                # После создания структуры, перезагружаем метаданные
                self.current_project_info = self.storage.load_project_metadata()
                return True
            else:
                self.logger.error("ProjectManager не смог инициализировать проект.")
                return False
        except Exception as e:
            self.logger.error(f"Ошибка при инициализации проекта через ProjectManager: {e}", exc_info=True)
            return False

    def load_project(self, project_path: str) -> bool:
        """
        Загрузка существующего проекта.

        Args:
            project_path (str): Путь к директории проекта

        Returns:
            bool: True если проект загружен успешно, False в противном случае
        """
        # if not self.is_initialized:
        #     logger.warning("Приложение не инициализировано. Выполняем инициализацию...")
        #     if not self.initialize():
        #         return False

        try:
            self.logger.info(f"Загрузка проекта из: {project_path}")

            if not self.storage:
                self.logger.error("Storage не инициализирован")
                return False

            project_data = self.storage.load_project_metadata() # <-- Загрузка через storage
            if project_data:
                self.current_project_info = project_data
                self.project_path = Path(project_path)
                self.logger.info("Проект загружен успешно")
                return True
            else:
                self.logger.error("Не удалось загрузить метаданные проекта")
                return False

        except Exception as e:
            self.logger.error(f"Ошибка при загрузке проекта: {e}", exc_info=True)
            return False

    def close_project(self):
        """Закрывает текущий проект и освобождает ресурсы."""
        self.logger.info("Закрытие проекта.")
        # Storage сам управляет соединением через контекстный менеджер или явный disconnect
        # if self.storage:
        #     self.storage.disconnect()
        self.storage = None
        self.current_project_info = None
        self.is_initialized = False
        self.logger.debug("Проект закрыт.")

    # --- Анализ Excel-файлов ---

    def analyze_file(self, file_path: str) -> bool:
        """
        Анализирует Excel-файл и сохраняет результаты в БД проекта.

        Args:
            file_path (str): Путь к анализируемому .xlsx файлу.

        Returns:
            bool: True, если анализ и сохранение успешны, иначе False.
        """
        if not self.is_initialized:
            self.logger.warning("AppController не инициализирован. Невозможно выполнить анализ")
            return False

        if not self.project_path:
            self.logger.error("Путь к проекту не установлен.")
            return False

        try:
            self.logger.info(f"Начало анализа файла: {file_path}")

            # Проверка существования файла
            if not Path(file_path).exists():
                self.logger.error(f"Файл не найден: {file_path}")
                return False

            # - ИНТЕГРАЦИЯ ANALYZER: Вызов анализатора -
            # Передаем путь к файлу и, возможно, storage для прямого сохранения (или возвращаем данные)
            # Более предпочтительно возвращать данные, а сохранение делегировать storage.
            documentation_data = analyze_workbook_logic(file_path) # <-- Новое имя функции
            if documentation_data is None:
                self.logger.error("Анализатор вернул None. Ошибка при анализе файла.")
                return False

            self.logger.info("Анализ файла завершен успешно")

            # --- НАЧАЛО ИНТЕГРАЦИИ СО ХРАНИЛИЩЕМ ---
            # Сохраняем результаты анализа в БД через self.storage
            # Предположим, analyze_workbook_logic возвращает структуру вроде:
            # {'project_name': '...', 'sheets': [{'name': 'Sheet1', 'raw_data': [...], 'formulas': [...], ...}]}
            project_name = documentation_data.get("project_name", self.project_path.name if self.project_path else "UnknownProject")
            sheets_data = documentation_data.get("sheets", [])

            for sheet_info in sheets_data:
                sheet_name = sheet_info.get("name")
                if not sheet_name:
                    self.logger.warning("Имя листа отсутствует в данных анализа. Пропуск.")
                    continue

                self.logger.info(f"Сохранение данных для листа: {sheet_name}")

                # --- Получаем или создаем запись листа в БД через storage ---
                # storage теперь должен уметь создавать листы, если их нет
                # TODO: Реализовать метод в storage для создания/получения sheet_id
                # Пока используем вспомогательный метод, но вызываем storage методы напрямую
                # Получаем sheet_id через storage, если он может его вычислить или создать запись
                # storage.create_or_get_sheet(project_name, sheet_name) -> sheet_id
                # Пока предположим, что storage может работать с именем листа напрямую в большинстве случаев

                # --- Сохраняем метаданные листа ---
                metadata_to_save = {
                    "max_row": sheet_info.get("max_row"),
                    "max_column": sheet_info.get("max_column"),
                    "merged_cells": sheet_info.get("merged_cells", []),
                    # Добавьте другие метаданные, если есть
                }
                # storage.save_sheet_metadata(sheet_name, metadata_to_save) # <-- Вызов метода storage

                # --- Сохраняем "сырые данные" ---
                raw_data_to_save = sheet_info.get("raw_data", [])
                # storage.save_sheet_raw_data(sheet_name, raw_data_to_save) # <-- Вызов метода storage

                # --- Сохраняем редактируемые данные ---
                editable_data_to_save = sheet_info.get("editable_data", []) # Если анализатор возвращает
                # storage.save_sheet_editable_data(sheet_name, editable_data_to_save) # <-- Вызов метода storage

                # --- Сохраняем формулы ---
                formulas_to_save = sheet_info.get("formulas", [])
                # storage.save_sheet_formulas(sheet_name, formulas_to_save) # <-- Вызов метода storage

                # --- Сохраняем стили ---
                styles_to_save = sheet_info.get("styles", [])
                # storage.save_sheet_styles(sheet_name, styles_to_save) # <-- Вызов метода storage

                # --- Сохраняем диаграммы ---
                charts_to_save = sheet_info.get("charts", [])
                # storage.save_sheet_charts(sheet_name, charts_to_save) # <-- Вызов метода storage

                # --- Используем методы из модулей storage ---
                # raw_data
                from src.storage.raw_data import save_sheet_raw_data
                save_sheet_raw_data(self.storage.connection, sheet_name, raw_data_to_save)
                # editable_data
                from src.storage.editable_data import save_sheet_editable_data
                save_sheet_editable_data(self.storage.connection, sheet_name, editable_data_to_save)
                # formulas
                from src.storage.formulas import save_sheet_formulas
                save_sheet_formulas(self.storage.connection, sheet_name, formulas_to_save)
                # styles
                from src.storage.styles import save_sheet_styles
                save_sheet_styles(self.storage.connection, sheet_name, styles_to_save)
                # charts
                from src.storage.charts import save_sheet_charts
                save_sheet_charts(self.storage.connection, sheet_name, charts_to_save)
                # metadata
                from src.storage.metadata import save_sheet_metadata
                save_sheet_metadata(self.storage.connection, sheet_name, metadata_to_save)

            self.logger.info(f"Анализ и сохранение данных из '{file_path}' завершены.")
            return True

        except Exception as e:
            self.logger.error(f"Ошибка при анализе/сохранении файла '{file_path}': {e}", exc_info=True)
            return False


    # --- Работа с данными листа (для GUI) ---

    def get_sheet_data(self, sheet_name: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        Получает данные листа для отображения в GUI.
        Возвращает кортеж (raw_data, editable_data).

        Args:
            sheet_name (str): Имя листа.

        Returns:
            Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
            Кортеж из списков сырых и редактируемых данных.
        """
        if not self.is_initialized:
            self.logger.error("AppController не инициализирован.")
            return ([], [])

        try:
            self.logger.debug(f"Загрузка данных для листа '{sheet_name}'...")

            # Загружаем "сырые данные"
            # --- Используем метод из модуля storage ---
            from src.storage.raw_data import load_sheet_raw_data
            raw_data = load_sheet_raw_data(self.storage.connection, sheet_name)
            self.logger.debug(f"Загружено {len(raw_data)} записей сырых данных.")

            # Загружаем редактируемые данные
            # --- Используем метод из модуля storage ---
            from src.storage.editable_data import load_sheet_editable_data
            editable_data_result = load_sheet_editable_data(self.storage.connection, sheet_name)
            editable_data = editable_data_result.get('rows', []) if editable_data_result else []
            self.logger.debug(f"Загружено {len(editable_data)} записей редактируемых данных.")

            return (raw_data, editable_data)

        except Exception as e:
            self.logger.error(f"Ошибка при загрузке данных для листа '{sheet_name}': {e}", exc_info=True)
            return ([], [])

    def update_cell_value(self, sheet_name: str, cell_address: str, new_value: Any) -> bool:
        """
        Обновляет значение ячейки и записывает изменение в историю.
        (Адаптировано из update_sheet_cell_in_project)

        Args:
            sheet_name (str): Имя листа.
            cell_address (str): Адрес ячейки (например, 'A1').
            new_value (Any): Новое значение ячейки.

        Returns:
            bool: True, если обновление успешно, иначе False.
        """
        if not self.is_initialized:
            self.logger.warning("AppController не инициализирован. Невозможно обновить ячейку.")
            return False

        try:
            self.logger.debug(f"Обновление ячейки {cell_address} на листе '{sheet_name}' на значение '{new_value}'.")

            # 1. Получаем sheet_id через storage (или напрямую из таблицы, если имя достаточно)
            # Используем вспомогательный метод, адаптированный к новому storage
            sheet_id = self._get_sheet_id_by_name(sheet_name)
            if sheet_id is None:
                self.logger.error(f"Не найден sheet_id для листа '{sheet_name}'. Обновление невозможно.")
                return False

            # 2. Получаем старое значение (для истории)
            # Это может быть сложно, так как оно может быть в raw_data или editable_data
            # Для простоты MVP, предположим, что мы можем получить его из editable_data
            # или установить None. Более точная логика может потребоваться.
            # Получим через editable_data.load, найдем ячейку
            from src.storage.editable_data import load_sheet_editable_data
            current_editable_data = load_sheet_editable_data(self.storage.connection, sheet_name)
            old_value = None
            if current_editable_data and 'column_names' in current_editable_data and 'rows' in current_editable_data:
                 col_names = current_editable_data['column_names']
                 rows = current_editable_data['rows']
                 # Преобразуем A1 в (row_idx, col_name)
                 import re
                 match = re.match(r"([A-Z]+)(\d+)", cell_address)
                 if match:
                     col_letter, row_num = match.groups()
                     row_idx = int(row_num) - 1 # 0-based
                     # Преобразуем букву столбца в индекс
                     col_idx = 0
                     for c in col_letter:
                         col_idx = col_idx * 26 + (ord(c) - ord('A') + 1)
                     col_idx -= 1 # 0-based
                     if 0 <= row_idx < len(rows) and 0 <= col_idx < len(col_names):
                         old_value = rows[row_idx][col_idx]

            # 3. Обновляем редактируемые данные
            # --- Используем метод из модуля storage ---
            from src.storage.editable_data import update_editable_cell
            update_success = update_editable_cell(self.storage.connection, sheet_name, row_idx, col_names[col_idx], new_value)

            if update_success:
                # 4. Записываем в историю редактирования
                # --- Используем метод из модуля storage.history ---
                from src.storage.history import save_edit_history_record
                # Требуется project_id, sheet_id, cell_address, action_type, old_value, new_value, user, details
                # project_id пока получим из метаданных или предположим 1
                project_id = 1 # TODO: Получить реальный project_id из БД
                 # details
                history_details = {
                    "cell_address": cell_address,
                    "row_index": row_idx,
                    "column_index": col_idx
                }
                history_success = save_edit_history_record(
                     connection=self.storage.connection,
                     project_id=project_id,
                     sheet_id=sheet_id,
                     cell_address=cell_address,
                     action_type="edit_cell",
                     old_value=old_value,
                     new_value=new_value,
                     user=None, # TODO: Поддержка пользователей
                     details=history_details
                )
                if history_success:
                    self.logger.info(f"Ячейка {cell_address} на листе '{sheet_name}' успешно обновлена и история сохранена.")
                    return True
                else:
                    self.logger.error(
                            f"Ячейка обновлена, но ошибка при сохранении записи истории для {cell_address} на листе '{sheet_name}'.")
                    # Возвращаем False, если история критична
                    return False
            else:
                self.logger.error(f"Не удалось обновить ячейку {cell_address} на листе '{sheet_name}' в БД.")
                return False

        except Exception as e:
            self.logger.error(f"Ошибка при обновлении ячейки {cell_address} на листе '{sheet_name}': {e}", exc_info=True)
            return False

    def _get_sheet_id_by_name(self, sheet_name: str) -> Optional[int]:
        """Вспомогательный метод для получения sheet_id по имени листа."""
        if not self.storage or not self.storage.connection:
             return None
        try:
            cursor = self.storage.connection.cursor()
            # Предполагаем project_id = 1
            cursor.execute("SELECT sheet_id FROM sheets WHERE name = ? AND project_id = 1", (sheet_name,))
            result = cursor.fetchone()
            return result[0] if result else None
        except sqlite3.Error:
             return None

    # --- Экспорт ---

    def export_project(self, output_path: str) -> bool:
        """
        Экспортирует проект в Excel-файл.

        Args:
            output_path (str): Путь к выходному .xlsx файлу.

        Returns:
            bool: True, если экспорт успешен, иначе False.
        """
        if not self.is_initialized:
            self.logger.warning("AppController не инициализирован. Невозможно выполнить экспорт.")
            return False

        if not self.project_path:
            self.logger.error("Путь к проекту не установлен.")
            return False

        try:
            self.logger.info(f"Начало экспорта проекта в Excel: {output_path}")

            # --- ИНТЕГРАЦИЯ EXPORTER ---
            # Используем direct_db_exporter, передавая ему путь к БД проекта и путь к выходному файлу
            # Он прочитает все данные из БД через ProjectDBStorage и создаст Excel-файл

            # Определяем путь к БД проекта
            db_path = self.project_path / "project_data.db"

            success = export_project_db_to_excel(
                db_path=str(db_path),
                output_path=output_path
            )

            if success:
                self.logger.info("Экспорт в Excel завершен успешно.")
                return True
            else:
                self.logger.error("Ошибка при экспорте в Excel.")
                return False

        except Exception as e:
            self.logger.error(f"Ошибка при экспорте проекта в '{output_path}': {e}", exc_info=True)
            return False

    # --- История редактирования (Undo/Redo) ---

    def get_edit_history(self, sheet_name: Optional[str] = None, limit: Optional[int] = 10) -> List[Dict[str, Any]]:
        """
        Получает историю редактирования.

        Args:
            sheet_name (Optional[str]): Имя листа для фильтрации. Если None, вся история.
            limit (Optional[int]): Максимальное количество записей.

        Returns:
            List[Dict[str, Any]]: Список записей истории.
        """
        if not self.is_initialized:
            self.logger.error("AppController не инициализирован.")
            return []

        try:
             sheet_id = self._get_sheet_id_by_name(sheet_name) if sheet_name else None
             # Используем метод из модуля storage.history
             from src.storage.history import load_edit_history
             return load_edit_history(self.storage.connection, sheet_id, limit)
        except Exception as e:
            self.logger.error(f"Ошибка при загрузке истории редактирования: {e}", exc_info=True)
            return []

    # --- Обработка данных (заглушка) ---
    def process_data(self):
        """Заглушка для обработки данных."""
        self.logger.info("Обработка данных (заглушка)")
        # TODO: Интеграция с processor модулем
        pass

    # --- Другие методы ---
    def get_project_info(self) -> Optional[Dict[str, Any]]:
        """
        Получение информации о текущем проекте.

        Returns:
            Optional[Dict[str, Any]]: Информация о проекте или None если проект не загружен
        """
        if not self.is_initialized or not self.current_project_info:
            self.logger.warning("Проект не загружен")
            return None
        return self.current_project_info

    def shutdown(self) -> None:
        """
        Корректное завершение работы приложения.
        """
        self.logger.info("Завершение работы приложения")
        self.close_project() # Закрываем проект, если он был загружен
        self.logger.info("Приложение завершено")


# --- ФАБРИЧНЫЕ ФУНКЦИИ ---

def get_app_controller(project_path: str, log_level: str = "INFO") -> AppController:
    """Фабричная функция для создания и настройки AppController."""
    # Создаём ProjectManager и получаем путь к БД проекта
    project_manager = ProjectManager(log_level=log_level)
    project_data_path = project_manager.get_project_db_path(project_path)
    storage = ProjectDBStorage(project_data_path, log_level=log_level)
    controller = AppController(project_path=project_path, storage=storage, log_level=log_level)
    controller.initialize() # <-- ВАЖНО: инициализирует перед возвратом
    return controller

def create_app_controller(project_path: str, log_level: str = "INFO") -> AppController:
    """
    Фабричная функция для создания и настройки AppController.
    (Алиас для get_app_controller, возвращает полностью инициализированный контроллер).
    """
    # Просто вызываем get_app_controller, чтобы использовать общую логику
    return get_app_controller(project_path, log_level)

# --- Пример использования (опционально) ---
# if __name__ == "__main__":
#     # Это просто для демонстрации, не будет выполняться при импорте
#     logger.info("Демонстрация работы AppController")
#
#     # Создание контроллера через фабричную функцию
#     app_ctrl = get_app_controller("./test_project")
#
#     if app_ctrl.is_initialized:
#         logger.info("Контроллер приложения инициализирован успешно")
#         # Дальнейшие действия с app_ctrl
#         app_ctrl.shutdown()
#     else:
#         logger.error("Ошибка инициализации контроллера приложения")