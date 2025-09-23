# src/core/app_controller.py
import os
import sys
from pathlib import Path
import sqlite3
from typing import Dict, Any, List, Optional, Tuple

# --- ИНТЕГРАЦИЯ: Импорты компонентов ---
# from src.storage.database import ProjectDBStorage # <-- СТАРОЕ
from src.storage.base import ProjectDBStorage    # <-- НОВОЕ
from src.analyzer.logic_documentation import analyze_excel_file # <-- Правильное имя функции
# from src.exporter.excel_exporter import export_project as export_with_xlsxwriter # <-- Потенциальная проблема с xlsxwriter
# from src.exporter.direct_db_exporter import export_project as export_with_openpyxl # <-- Потенциальная проблема с именем функции
# from src.exporter.direct_db_exporter import export_project_db_to_excel # <-- Проверим настоящее имя
# from src.utils.logger import setup_logger # <-- Проверим настоящее имя
# from src.core.project_manager import ProjectManager
# from src.exceptions.app_exceptions import (...) # <-- Проверим настоящие имена

# Получаем логгер для этого модуля
# logger = setup_logger(__name__)
from src.utils.logger import get_logger # <-- Используем get_logger из старого main.py
logger = get_logger(__name__)

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
        # self.logger = setup_logger(self.__class__.__name__, log_level=log_level) # <-- setup_logger не найден
        from src.utils.logger import get_logger # <-- Временное решение для логгера
        self.logger = get_logger(self.__class__.__name__)
        self.current_project_info = None # <-- Информация о метаданных проекта
        self.is_initialized = False # <-- Статус инициализации AppController
        # --- СООТВЕТСТВИЕ СТАРОМУ main.py ---
        # Добавим атрибут, который ожидает старый main.py
        self.is_project_loaded = False # <-- Используется в старом main.py
        logger.debug("AppController инициализирован")

    def initialize(self):
        """Инициализация внутреннего состояния AppController."""
        try:
            self.logger.info("Инициализация внутреннего состояния AppController...")
            if not self.storage:
                 self.logger.error("Storage не предоставлен для инициализации AppController.")
                 return False
            # Инициализируем схему БД, если нужно
            # self.storage.initialize_schema() # <-- initialize_schema не найден в base.py напрямую
            self.storage.initialize_schema() # <-- Проверим, как это реализовано в base.py
            # NOTE: В base.py initialize_schema() вызывает connect(), а затем _create_tables()
            # Это должно установить self.connection
            # Пытаемся загрузить метаданные проекта
            # storage.load_project_metadata() вызывает метод в metadata.py через base.py
            # self.current_project_info = self.storage.load_project_metadata() # <-- load_project_metadata не найден напрямую
            # NOTE: В base.py load_project_metadata() вызывает _load_project_metadata из metadata.py
            self.current_project_info = self.storage.load_project_metadata() # <-- Проверим, как это реализовано в base.py
            if self.current_project_info:
                self.logger.info(f"Метаданные проекта загружены: {self.current_project_info.get('name', 'Unknown')}")
                # --- СООТВЕТСТВИЕ СТАРОМУ main.py ---
                self.is_project_loaded = True # <-- Установим флаг, как в старом AppController
            else:
                self.logger.info("Метаданные проекта не найдены, будет создан новый проект.")
                # --- СООТВЕТСТВИЕ СТАРОМУ main.py ---
                # self.is_project_loaded останется False, как ожидалось

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
            # project_manager = ProjectManager(log_level=self.log_level) # <-- log_level не найден
            project_manager = ProjectManager() # <-- Проверим конструктор PM
            # --- ИСПРАВЛЕНО: Вызов правильного метода ---
            # success = project_manager.create_project_structure(str(self.project_path), self.storage)
            # success = project_manager.create_project(project_path=str(self.project_path), storage=self.storage) # <-- storage не найден
            success = project_manager.create_project(str(self.project_path), self.storage) # <-- Проверим сигнатуру
            # ---
            if success:
                self.logger.info("Проект успешно инициализирован через ProjectManager.")
                # После создания структуры, перезагружаем метаданные
                # self.current_project_info = self.storage.load_project_metadata() # <-- load_project_metadata не найден
                self.current_project_info = self.storage.load_project_metadata() # <-- Проверим
                # --- СООТВЕТСТВИЕ СТАРОМУ main.py ---
                self.is_project_loaded = True # <-- Установим флаг после успешного создания
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
        try:
            self.logger.info(f"Загрузка проекта из: {project_path}")

            if not self.storage:
                self.logger.error("Storage не инициализирован")
                return False

            # project_data = self.storage.load_project_metadata() # <-- Загрузка через storage
            # NOTE: load_project_metadata загружает метаданные ТЕКУЩЕГО проекта (установленного в storage при инициализации)
            # Для загрузки другого проекта нужно пересоздать storage или передать путь явно.
            # В текущей архитектуре ProjectManager управляет созданием/загрузкой проекта на уровне файлов/папок и БД.
            # AppController работает с уже открытым проектом через storage.
            # Поэтому load_project тут бессмысленен, если только не пересоздаётся storage.
            # Пока оставим как есть, но флаг is_project_loaded установим вручную.
            # project_data = self.storage.load_project_metadata() # <-- Загрузка через storage
            # if project_data:
            #     self.current_project_info = project_data
            #     self.project_path = Path(project_path)
            #     self.logger.info("Проект загружен успешно")
            #     self.is_project_loaded = True # <-- Установим флаг
            #     return True
            # else:
            #     self.logger.error("Не удалось загрузить метаданные проекта")
            #     return False

            # Учитывая архитектуру, load_project может быть задачей ProjectManager или AppController.create_app_controller/get_app_controller
            # Просто установим путь и флаг, если storage уже готов.
            # Это не полноценная загрузка, но соответствует ожиданиям старого main.py.
            self.project_path = Path(project_path)
            # Предполагаем, что storage уже указывает на правильную БД для этого пути (через ProjectManager)
            # и что инициализация (которая загружает метаданные) уже была выполнена.
            if self.is_initialized: # Если уже инициализирован с нужной БД
                 self.is_project_loaded = True
                 self.logger.info("Проект (метаданные) считается загруженным через инициализацию.")
                 return True
            else:
                 self.logger.error("AppController не инициализирован, невозможно загрузить проект.")
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
        # --- СООТВЕТСТВИЕ СТАРОМУ main.py ---
        self.is_project_loaded = False # <-- Сбросим флаг
        self.logger.debug("Проект закрыт.")

    # --- Анализ Excel-файлов ---

    # --- ИСПРАВЛЕНО: Переименовано для соответствия main.py ---
    def analyze_excel_file(self, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
    # ---
        """
        Анализирует Excel-файл и сохраняет результаты в БД проекта.

        Args:
            file_path (str): Путь к анализируемому .xlsx файлу.
            options (Optional[Dict[str, Any]]): Дополнительные опции анализа (не используется напрямую в этой версии)

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
            # NOTE: analyze_excel_file возвращает сырые данные, которые нужно сохранить через storage
            # NOTE 2: analyze_excel_file ожидает connection, а не storage. Нужно передать connection.
            # documentation_data = analyze_excel_file(file_path) # <-- Правильное имя функции, но неправильный аргумент
            documentation_data = analyze_excel_file(file_path, self.storage.connection) # <-- Передаём connection
            if documentation_data is None:
                self.logger.error("Анализатор вернул None. Ошибка при анализе файла.")
                return False

            self.logger.info("Анализ файла завершен успешно")

            # NOTE: В новой архитектуре analyze_excel_file (в своей новой версии) может сразу сохранять в БД.
            # Но если он возвращает данные, мы можем их обработать здесь.
            # Предположим, что analyze_excel_file уже сохранил, и нам не нужно сохранять снова.
            # Однако, в текущей реализации analyze_excel_file возвращает данные.
            # Проверим, что он делает. Если он сохраняет, то повторное сохранение не нужно.
            # Если он возвращает, то нужно сохранить.
            # В текущей версии analyze_excel_file возвращает словарь и НЕ сохраняет в БД.
            # AppController должен его обработать.
            # project_name = documentation_data.get("project_name", self.project_path.name if self.project_path else "UnknownProject")
            # sheets_data = documentation_data.get("sheets", [])

            # NOTE: ProjectDBStorage.base.py теперь координирует вызовы подмодулей.
            # Мы передаём ему всю структуру documentation_data, и он сам решает, как распределить по таблицам.
            # Однако, текущая реализация analyze_excel_file и storage.save_analysis_results может не совпадать.
            # Лучше вызывать методы подмодулей storage напрямую из AppController, как в старом коде.

            # --- СОХРАНЕНИЕ ЧЕРЕЗ ПОДМОДУЛИ STORAGE (аналогично предыдущей версии) ---
            # NOTE: ProjectDBStorage.base.py предоставляет доступ к connection и вызывает методы подмодулей.
            # Мы можем использовать self.storage.connection и вызывать функции из src.storage.* напрямую.
            # Это требует импорта этих функций.

            # NOTE: Поскольку analyze_excel_file НЕ сохраняет, мы должны сохранить данные здесь.
            project_name = documentation_data.get("project_name", self.project_path.name if self.project_path else "UnknownProject")
            sheets_data = documentation_data.get("sheets", [])

            for sheet_info in sheets_data:
                sheet_name = sheet_info.get("name")
                if not sheet_name:
                    self.logger.warning("Имя листа отсутствует в данных анализа. Пропуск.")
                    continue

                self.logger.info(f"Сохранение данных для листа: {sheet_name}")

                # raw_data
                from src.storage.raw_data import save_sheet_raw_data
                raw_data_to_save = sheet_info.get("raw_data", [])
                save_sheet_raw_data(self.storage.connection, sheet_name, raw_data_to_save)
                # editable_data
                from src.storage.editable_data import save_sheet_editable_data
                editable_data_to_save = sheet_info.get("editable_data", [])
                save_sheet_editable_data(self.storage.connection, sheet_name, editable_data_to_save)
                # formulas
                from src.storage.formulas import save_sheet_formulas
                formulas_to_save = sheet_info.get("formulas", [])
                save_sheet_formulas(self.storage.connection, sheet_name, formulas_to_save)
                # styles
                from src.storage.styles import save_sheet_styles
                styles_to_save = sheet_info.get("styles", [])
                save_sheet_styles(self.storage.connection, sheet_name, styles_to_save)
                # charts
                from src.storage.charts import save_sheet_charts
                charts_to_save = sheet_info.get("charts", [])
                save_sheet_charts(self.storage.connection, sheet_name, charts_to_save)
                # metadata
                from src.storage.metadata import save_sheet_metadata
                metadata_to_save = {
                    "max_row": sheet_info.get("max_row"),
                    "max_column": sheet_info.get("max_column"),
                    "merged_cells": sheet_info.get("merged_cells", []),
                }
                save_sheet_metadata(self.storage.connection, sheet_name, metadata_to_save)
                # misc (если есть)
                # from src.storage.misc import save_sheet_misc_data # <-- Если такой модуль есть
                # misc_to_save = sheet_info.get("misc", {})
                # save_sheet_misc_data(self.storage.connection, sheet_name, misc_to_save)


            self.logger.info(f"Анализ и сохранение данных из '{file_path}' завершены.")
            return True

        except Exception as e:
            self.logger.error(f"Ошибка при анализе/сохранении файла '{file_path}': {e}", exc_info=True)
            return False


    # --- Работа с данными листа (для GUI / старый main.py) ---

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
            # NOTE: load_sheet_raw_data ожидает connection
            raw_data = load_sheet_raw_data(self.storage.connection, sheet_name)
            self.logger.debug(f"Загружено {len(raw_data)} записей сырых данных.")

            # Загружаем редактируемые данные
            # --- Используем метод из модуля storage ---
            from src.storage.editable_data import load_sheet_editable_data
            # NOTE: load_sheet_editable_data ожидает connection, возвращает {'column_names': [...], 'rows': [...]}
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
        Использует логику, похожую на старую функцию, но с новыми вызовами storage.

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
            # Для простоты MVP, получим его из editable_data
            # Получим через editable_data.load, найдем ячейку
            from src.storage.editable_data import load_sheet_editable_data
            current_editable_data_result = load_sheet_editable_data(self.storage.connection, sheet_name)
            old_value = None
            row_idx = None
            col_idx = None
            col_names = None
            if current_editable_data_result and 'column_names' in current_editable_data_result and 'rows' in current_editable_data_result:
                 col_names = current_editable_data_result['column_names']
                 rows = current_editable_data_result['rows']
                 # Преобразуем A1 в (row_idx, col_name)
                 import re
                 match = re.match(r"([A-Z]+)(\d+)", cell_address)
                 if match:
                     col_letter, row_num_str = match.groups()
                     row_num = int(row_num_str)
                     row_idx = row_num - 1 # 0-based
                     # Преобразуем букву столбца в индекс
                     col_idx = 0
                     for c in col_letter:
                         col_idx = col_idx * 26 + (ord(c) - ord('A') + 1)
                     col_idx -= 1 # 0-based
                     if 0 <= row_idx < len(rows) and 0 <= col_idx < len(col_names):
                         old_value = rows[row_idx][col_idx]
                     else:
                         self.logger.error(f"Адрес ячейки {cell_address} выходит за пределы данных листа '{sheet_name}'.")
                         return False
                 else:
                     self.logger.error(f"Неверный формат адреса ячейки: {cell_address}")
                     return False
            else:
                 self.logger.error(f"Не удалось загрузить редактируемые данные для листа '{sheet_name}' для получения старого значения.")
                 return False # Или можно обновить, считая старое значение None?

            # 3. Обновляем редактируемые данные
            # --- Используем метод из модуля storage ---
            # NOTE: update_editable_cell ожидает connection, sheet_name, row_index (0-based), column_name, new_value
            # column_name - это имя столбца (например, 'A', 'B'), а не индекс.
            # col_names[col_idx] даст нам имя столбца.
            column_name_to_update = col_names[col_idx] if col_names and 0 <= col_idx < len(col_names) else f"Col_{col_idx}"
            from src.storage.editable_data import update_editable_cell
            # NOTE: update_editable_cell использует индекс строки (0-based) и имя столбца
            update_success = update_editable_cell(self.storage.connection, sheet_name, row_idx, column_name_to_update, new_value)

            if update_success:
                # 4. Записываем в историю редактирования
                # --- Используем метод из модуля storage.history ---
                from src.storage.history import save_edit_history_record
                # Требуется connection, project_id, sheet_id, cell_address, action_type, old_value, new_value, user, details
                # project_id пока получим из метаданных или предположим 1, если не удастся иначе
                # project_id = 1 # TODO: Получить реальный project_id из БД или storage
                # Получим project_id из current_project_info, если оно есть и содержит ID
                project_id = self.current_project_info.get('id') if self.current_project_info and 'id' in self.current_project_info else 1
                 # details
                history_details = {
                    "cell_address": cell_address,
                    "row_index": row_idx,
                    "column_index": col_idx,
                    "column_name": column_name_to_update
                }
                # --- ИСПРАВЛЕНО: Вызов с правильными параметрами ---
                # history_success = save_edit_history_record(
                #      connection=self.storage.connection,
                #      project_id=project_id,
                #      sheet_id=sheet_id,
                #      cell_address=cell_address, # или None, если используется row/col
                #      action_type="edit_cell", # <-- action_type не найден
                #      old_value=old_value,
                #      new_value=new_value,
                #      user=None, # TODO: Поддержка пользователей # <-- user не найден
                #      details=history_details # <-- details не найден
                # )
                history_success = save_edit_history_record(
                     connection=self.storage.connection,
                     project_id=project_id,
                     sheet_id=sheet_id,
                     cell_address=cell_address, # или None, если используется row/col
                     # action_type="edit_cell", # <-- УБРАНО
                     old_value=old_value,
                     new_value=new_value,
                     # user=None, # TODO: Поддержка пользователей # <-- УБРАНО
                     # details=history_details # <-- УБРАНО
                )
                # ---
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
            # NOTE: Предполагаем, что project_id можно получить или он фиксирован.
            # Получим project_id из current_project_info, если возможно.
            project_id = self.current_project_info.get('id') if self.current_project_info and 'id' in self.current_project_info else 1

            cursor = self.storage.connection.cursor()
            cursor.execute("SELECT sheet_id FROM sheets WHERE name = ? AND project_id = ?", (sheet_name, project_id))
            result = cursor.fetchone()
            return result[0] if result else None
        except sqlite3.Error as e:
             self.logger.error(f"Ошибка SQLite при получении sheet_id: {e}")
             return None

    # --- Экспорт ---

    # --- ИСПРАВЛЕНО: Переименовано для соответствия main.py ---
    def export_results(self, export_type: str, output_path: str) -> bool:
    # ---
        """
        Экспортирует проект в Excel-файл.

        Args:
            export_type (str): Тип экспорта (например, 'excel'). Сейчас поддерживается только 'excel'.
            output_path (str): Путь к выходному .xlsx файлу.

        Returns:
            bool: True, если экспорт успешен, иначе False.
        """
        # Пока поддерживаем только экспорт в Excel
        if export_type.lower() != 'excel':
            self.logger.warning(f"Тип экспорта '{export_type}' пока не поддерживается. Поддерживается только 'excel'.")
            return False

        # На самом деле, export_results в main.py вызывает app_controller.export_results(export_type=export_type, output_path=output_path)
        # И export_project в AppController принимает output_path
        # Значит, export_results должен вызвать export_project
        return self.export_project(output_path)

    def export_project(self, output_path: str) -> bool:
        """
        Экспортирует проект в Excel-файл (реализация).

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

            # --- ИСПРАВЛЕНО: Вызов правильной функции ---
            # success = export_with_openpyxl(self.project_db_path, output_path) # <-- Старое имя/параметры
            # from src.exporter.direct_db_exporter import export_project_db_to_excel # <-- Проверим настоящее имя
            from src.exporter.direct_db_exporter import export_project_db_to_excel # <-- Предположим, что это правильное имя
            success = export_project_db_to_excel(
                db_path=str(db_path),
                output_path=output_path
            )
            # ---

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
             # NOTE: load_edit_history ожидает connection, sheet_id, limit
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
# NOTE: create_app_controller должна принимать project_path, создавать storage и возвращать инициализированный AppController

def get_app_controller(project_path: str, log_level: str = "INFO") -> AppController:
    """Фабричная функция для создания и настройки AppController."""
    # Создаём ProjectManager и получаем путь к БД проекта
    # project_manager = ProjectManager(log_level=log_level) # <-- log_level не найден у PM
    project_manager = ProjectManager() # <-- Используем конструктор без log_level
    # --- ИСПРАВЛЕНО: Вызов правильного метода ---
    # project_data_path = project_manager.get_project_db_path(project_path) # <-- Этого метода нет в PM
    # Предположим, что БД всегда называется project_data.db внутри project_path
    project_data_path = os.path.join(project_path, "project_data.db")
    # ---
    storage = ProjectDBStorage(project_data_path, log_level=log_level)
    controller = AppController(project_path=project_path, storage=storage, log_level=log_level)
    controller.initialize() # <-- ВАЖНО: инициализирует перед возвратом
    return controller

# --- ИСПРАВЛЕНО: create_app_controller теперь принимает project_path ---
def create_app_controller(project_path: str, log_level: str = "INFO") -> AppController:
    """
    Фабричная функция для создания и настройки AppController.
    (Алиас для get_app_controller, возвращает полностью инициализированный контроллер).
    """
    # Просто вызываем get_app_controller, чтобы использовать общую логику
    return get_app_controller(project_path, log_level)
# ---