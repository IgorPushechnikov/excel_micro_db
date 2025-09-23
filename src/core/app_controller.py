# src/core/app_controller.py

import os
import sqlite3
import logging
from typing import Dict, Any, List, Optional, Tuple

# Импортируем анализатор
from src.analyzer.logic_documentation import analyze_excel_file

# Импортируем хранилище
from src.storage.base import ProjectDBStorage

# Импортируем экспортёры
from src.exporter.excel_exporter import export_project as export_with_xlsxwriter # Основной
from src.exporter.direct_db_exporter import export_project as export_with_openpyxl # Аварийный (убедитесь, что она там есть)

# Импортируем logger из utils
from src.utils.logger import get_logger

# --- Исключения ---
from src.exceptions.app_exceptions import ProjectError, AnalysisError, ExportError

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
        self._current_project_data: Optional[Dict[str, Any]] = None # Кэш метаданных проекта
        logger.debug(f"AppController инициализирован для проекта: {project_path}")

    # --- Управление проектом ---

    def create_new_project(self, project_name: str) -> bool:
        """
        Создает новую структуру проекта и инициализирует БД.

        Args:
            project_name (str): Имя нового проекта.

        Returns:
            bool: True, если проект создан успешно, иначе False.
        """
        logger.info(f"Создание нового проекта: {project_name}")
        try:
            # Создаем директорию проекта, если её нет
            os.makedirs(self.project_path, exist_ok=True)
            logger.debug(f"Директория проекта создана/проверена: {self.project_path}")

            # Инициализируем хранилище и схему БД
            self.storage = ProjectDBStorage(self.project_db_path)
            if not self.storage.initialize_project_tables():
                logger.error("Не удалось инициализировать схему БД проекта.")
                return False
            
            # TODO: Сохранить метаданные проекта (имя, дата создания) в БД
            # Это может быть сделано через storage или напрямую
            logger.info(f"Проект '{project_name}' успешно создан в {self.project_path}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при создании проекта '{project_name}': {e}", exc_info=True)
            return False

    def load_project(self) -> bool:
        """
        Загружает существующий проект.

        Returns:
            bool: True, если проект загружен успешно, иначе False.
        """
        logger.info(f"Загрузка проекта из: {self.project_path}")
        if not os.path.exists(self.project_db_path):
            logger.error(f"Файл БД проекта не найден: {self.project_db_path}")
            return False

        try:
            self.storage = ProjectDBStorage(self.project_db_path)
            if not self.storage.connect():
                logger.error("Не удалось подключиться к БД проекта.")
                return False
            
            # TODO: Загрузить метаданные проекта в self._current_project_data
            logger.info("Проект успешно загружен.")
            return True
        except Exception as e:
            logger.error(f"Ошибка при загрузке проекта из {self.project_path}: {e}", exc_info=True)
            return False

    def close_project(self):
        """Закрывает текущий проект и освобождает ресурсы."""
        logger.info("Закрытие проекта.")
        if self.storage:
            self.storage.disconnect()
        self.storage = None
        self._current_project_data = None
        logger.debug("Проект закрыт.")

    # --- Анализ Excel-файлов ---

    def analyze_excel_file(self, file_path: str) -> bool:
        """
        Анализирует Excel-файл и сохраняет результаты в БД проекта.

        Args:
            file_path (str): Путь к анализируемому .xlsx файлу.

        Returns:
            bool: True, если анализ и сохранение успешны, иначе False.
        """
        if not self.storage:
            logger.error("Проект не загружен. Невозможно выполнить анализ.")
            return False

        if not os.path.exists(file_path):
            logger.error(f"Excel-файл для анализа не найден: {file_path}")
            return False

        logger.info(f"Начало анализа Excel-файла: {file_path}")
        try:
            # 1. Анализ файла
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
        if not self.storage:
            logger.error("Проект не загружен.")
            return ([], [])

        try:
            logger.debug(f"Загрузка данных для листа '{sheet_name}'...")
            
            # Загружаем "сырые данные"
            raw_data = self.storage.load_sheet_raw_data(sheet_name)
            logger.debug(f"Загружено {len(raw_data)} записей сырых данных.")

            # Получаем sheet_id для загрузки редактируемых данных
            sheet_id = self._get_sheet_id_by_name(sheet_name)
            if sheet_id is None:
                 logger.warning(f"Не найден sheet_id для листа '{sheet_name}'. Редактируемые данные не загружены.")
                 editable_data = []
            else:
                # Загружаем редактируемые данные
                # --- Используем исправленный метод из storage ---
                editable_data = self.storage.load_sheet_editable_data(sheet_id, sheet_name)
                logger.debug(f"Загружено {len(editable_data)} записей редактируемых данных.")
            
            return (raw_data, editable_data)

        except Exception as e:
            logger.error(f"Ошибка при загрузке данных для листа '{sheet_name}': {e}", exc_info=True)
            return ([], [])

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

    def update_cell_value(self, sheet_name: str, cell_address: str, new_value: Any) -> bool:
        """
        Обновляет значение ячейки и записывает изменение в историю.

        Args:
            sheet_name (str): Имя листа.
            cell_address (str): Адрес ячейки (например, 'A1').
            new_value (Any): Новое значение ячейки.

        Returns:
            bool: True, если обновление успешно, иначе False.
        """
        if not self.storage:
            logger.error("Проект не загружен. Невозможно обновить ячейку.")
            return False

        try:
            logger.debug(f"Обновление ячейки {cell_address} на листе '{sheet_name}' на значение '{new_value}'.")

            # 1. Получаем sheet_id
            sheet_id = self._get_sheet_id_by_name(sheet_name)
            if sheet_id is None:
                logger.error(f"Не найден sheet_id для листа '{sheet_name}'. Обновление невозможно.")
                return False

            # 2. Получаем старое значение (для истории)
            # Это может быть сложно, так как оно может быть в raw_data или editable_data
            # Для простоты MVP, предположим, что мы можем получить его из editable_data
            # или установить None. Более точная логика может потребоваться.
            old_value = None # TODO: Получить реальное старое значение
            
            # 3. Обновляем редактируемые данные
            # --- Используем исправленный метод из storage ---
            if not self.storage.update_editable_cell(sheet_id, sheet_name, cell_address, new_value):
                logger.error(f"Не удалось обновить редактируемую ячейку {cell_address} на листе '{sheet_name}'.")
                return False

            # 4. Записываем в историю редактирования
            # --- Используем метод из storage ---
            if not self.storage.save_edit_history_record(sheet_id, cell_address, old_value, new_value):
                 logger.warning(f"Не удалось записать изменение ячейки {cell_address} в историю.")

            logger.info(f"Ячейка {cell_address} на листе '{sheet_name}' успешно обновлена.")
            return True

        except Exception as e:
            logger.error(f"Ошибка при обновлении ячейки {cell_address} на листе '{sheet_name}': {e}", exc_info=True)
            return False

    # --- Экспорт ---

    def export_project(self, output_path: str, use_xlsxwriter: bool = True) -> bool:
        """
        Экспортирует проект в Excel-файл.

        Args:
            output_path (str): Путь к выходному .xlsx файлу.
            use_xlsxwriter (bool): Если True, использует основной экспортер (xlsxwriter).
                                   Если False, использует аварийный (openpyxl).

        Returns:
            bool: True, если экспорт успешен, иначе False.
        """
        logger.info(f"Начало экспорта проекта в '{output_path}'. Используется {'xlsxwriter' if use_xlsxwriter else 'openpyxl'}.")
        
        try:
            if use_xlsxwriter:
                # Пока используем заглушку, так как основной экспортер может быть еще не готов
                # или его путь импорта отличается.
                logger.warning("Основной экспортёр (xlsxwriter) пока не подключен напрямую. Используется аварийный.")
                success = export_with_openpyxl(self.project_db_path, output_path)
                # TODO: Подключить основной экспортер
                # from src.exporter.excel_exporter import export_project as export_with_xlsxwriter
                # success = export_with_xlsxwriter(self.project_db_path, output_path)
            else:
                success = export_with_openpyxl(self.project_db_path, output_path)
            
            if success:
                logger.info(f"Проект успешно экспортирован в '{output_path}'.")
            else:
                logger.error(f"Ошибка при экспорте проекта в '{output_path}'.")
            return success
            
        except Exception as e:
            logger.error(f"Неожиданная ошибка при экспорте проекта в '{output_path}': {e}", exc_info=True)
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
        if not self.storage:
            logger.error("Проект не загружен.")
            return []

        try:
             sheet_id = self._get_sheet_id_by_name(sheet_name) if sheet_name else None
             return self.storage.load_edit_history(sheet_id, limit)
        except Exception as e:
            logger.error(f"Ошибка при загрузке истории редактирования: {e}", exc_info=True)
            return []

    # Другие методы AppController (например, для processor, AppData) будут добавлены позже
