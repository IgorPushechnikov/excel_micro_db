# src/storage/base.py

import sqlite3
import logging
from contextlib import contextmanager
from typing import List, Dict, Any, Optional, Union
import os
import json

# Импортируем новые функции из модулей storage
from src.storage import schema, raw_data, editable_data, formulas, styles, charts, history, metadata, sheets

# Импортируем logger из utils
from src.utils.logger import get_logger

logger = get_logger(__name__)


class ProjectDBStorage:
    """
    Основной класс для взаимодействия с базой данных проекта SQLite.

    Координирует вызовы подмодулей для работы с различными аспектами данных проекта.
    """

    def __init__(self, db_path: str):
        """
        Инициализирует объект хранилища проекта.

        Args:
            db_path (str): Путь к файлу базы данных SQLite проекта.
        """
        self.db_path = db_path
        self.connection: Optional[sqlite3.Connection] = None
        logger.debug(f"ProjectDBStorage инициализирован с путем к БД: {db_path}")

    def connect(self) -> bool:
        """
        Устанавливает соединение с базой данных проекта.
        Если файл БД не существует, он будет создан при первом обращении к нему
        (например, в initialize_project_tables).

        Returns:
            bool: True, если соединение успешно установлено, иначе False.
        """
        try:
            # Убираем проверку существования файла.
            # SQLite создаст его при первом обращении (CREATE TABLE и т.д.)
            self.connection = sqlite3.connect(self.db_path)
            logger.info(f"Установлено соединение с БД проекта: {self.db_path}")
            return True
        except sqlite3.Error as e:
            logger.error(f"Ошибка подключения к БД проекта {self.db_path}: {e}")
            self.connection = None
            return False

    def disconnect(self):
        """Закрывает соединение с базой данных проекта."""
        if self.connection:
            try:
                self.connection.close()
                logger.info("Соединение с БД проекта закрыто.")
            except sqlite3.Error as e:
                logger.error(f"Ошибка при закрытии соединения с БД: {e}")
            finally:
                self.connection = None

    @contextmanager
    def get_connection(self):
        """
        Контекстный менеджер для автоматического управления соединением.
        """
        was_connected = self.connection is not None
        if not was_connected:
            # Пытаемся подключиться. Если файл не существует, connect() все равно вернет True,
            # а ошибка возникнет позже при выполнении запроса.
            if not self.connect():
                raise Exception(f"Не удалось подключиться к БД: {self.db_path}")
        try:
            yield self.connection
        finally:
            if not was_connected and self.connection:
                self.disconnect()

    def initialize_project_tables(self) -> bool:
        """
        Инициализирует схему таблиц проекта в БД.
        Создает новый файл БД, если он не существует.

        Returns:
            bool: True, если инициализация успешна, иначе False.
        """
        try:
            # Вместо использования get_connection (который вызывает connect и проверяет существование файла),
            # создаем соединение напрямую. Это позволит SQLite создать файл БД, если его нет.
            self.connection = sqlite3.connect(self.db_path)
            logger.info(f"Создано соединение с БД проекта (новый файл): {self.db_path}")
            
            # Теперь инициализируем схему
            schema.initialize_project_schema(self.connection)
            logger.info("Схема таблиц проекта инициализирована.")
            
            # Отключаемся после инициализации
            self.disconnect()
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при инициализации схемы таблиц проекта: {e}", exc_info=True)
            # Попытаемся отключиться, если соединение было установлено
            if self.connection:
                self.disconnect()
            return False

    # --- Методы для работы с листами (таблица sheets) ---

    def save_sheet(self, project_id: int, sheet_name: str, max_row: Optional[int] = None, max_column: Optional[int] = None) -> Optional[int]:
        """
        Сохраняет информацию о листе в таблицу 'sheets'.
        Если лист с таким именем для проекта уже существует, возвращает его sheet_id.
        Иначе создает новую запись.

        Args:
            project_id (int): ID проекта.
            sheet_name (str): Имя листа Excel.
            max_row (Optional[int]): Максимальный номер строки.
            max_column (Optional[int]): Максимальный номер столбца.

        Returns:
            Optional[int]: sheet_id листа или None в случае ошибки.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return sheets.save_sheet(conn, project_id, sheet_name, max_row, max_column)
                else:
                    return None
        except Exception as e:
            logger.error(f"Ошибка при сохранении листа '{sheet_name}': {e}", exc_info=True)
            return None

    def load_all_sheets_metadata(self, project_id: int = 1) -> List[Dict[str, Any]]:
        """
        Загружает метаданные (ID и имя) для всех листов в проекте.
        Используется для экспорта, чтобы знать, какие листы обрабатывать.

        Args:
            project_id (int): ID проекта (по умолчанию 1 для MVP).

        Returns:
            List[Dict[str, Any]]: Список словарей с ключами 'sheet_id' и 'name'.
            Возвращает пустой список в случае ошибки или отсутствия листов.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return sheets.load_all_sheets_metadata(conn, project_id)
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке списка листов для проекта ID {project_id}: {e}", exc_info=True)
            return []

    # --- Методы для работы с дополнительными метаданными листов (таблица project_metadata) ---

    def save_sheet_metadata(self, sheet_name: str, sheet_data: Dict[str, Any]) -> bool:
        """
        Сохраняет метаданные листа в БД проекта (в таблицу project_metadata).

        Args:
            sheet_name (str): Имя листа Excel.
            sheet_data (Dict[str, Any]): Словарь с метаданными листа (max_row, max_column и т.д.).

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return metadata.save_sheet_metadata(conn, sheet_name, sheet_data)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при сохранении метаданных листа '{sheet_name}': {e}", exc_info=True)
            return False

    def load_sheet_metadata(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """
        Загружает метаданные листа из БД проекта (из таблицы project_metadata).

        Args:
            sheet_name (str): Имя листа Excel.

        Returns:
            Optional[Dict[str, Any]]: Словарь с метаданными листа или None в случае ошибки.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return metadata.load_sheet_metadata(conn, sheet_name)
                else:
                    return None
        except Exception as e:
            logger.error(f"Ошибка при загрузке метаданных листа '{sheet_name}': {e}", exc_info=True)
            return None

    # --- Методы для работы с "сырыми" данными ---

    def save_sheet_raw_data(self, sheet_name: str, raw_data_list: List[Dict[str, Any]]) -> bool:
        """
        Сохраняет "сырые" данные листа в БД проекта.

        Args:
            sheet_name (str): Имя листа Excel.
            raw_data_list (List[Dict[str, Any]]): Список словарей с 'cell_address', 'value', 'value_type'.

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return raw_data.save_sheet_raw_data(conn, sheet_name, raw_data_list)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при сохранении сырых данных листа '{sheet_name}': {e}", exc_info=True)
            return False

    def load_sheet_raw_data(self, sheet_name: str) -> List[Dict[str, Any]]:
        """
        Загружает "сырые" данные листа из БД проекта.

        Args:
            sheet_name (str): Имя листа Excel.

        Returns:
            List[Dict[str, Any]]: Список словарей с 'cell_address', 'value', 'value_type'.
            Возвращает пустой список в случае ошибки или отсутствия данных.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return raw_data.load_sheet_raw_data(conn, sheet_name)
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке сырых данных листа '{sheet_name}': {e}", exc_info=True)
            return []

    # --- Методы для работы с редактируемыми данными ---

    # Используют функции из src/storage/editable_data.py

    def load_sheet_editable_data(self, sheet_id: int, sheet_name: str) -> List[Dict[str, Any]]:
        """
        Загружает редактируемые данные для указанного листа.

        Args:
            sheet_id (int): ID листа в БД.
            sheet_name (str): Имя листа Excel.

        Returns:
            List[Dict[str, Any]]: Список словарей с ключами 'cell_address' и 'value'.
            Возвращает пустой список в случае ошибки или отсутствия данных.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return editable_data.load_sheet_editable_data(conn, sheet_id, sheet_name)
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке редактируемых данных для листа '{sheet_name}' (ID: {sheet_id}): {e}", exc_info=True)
            return []

    def update_editable_cell(self, sheet_id: int, sheet_name: str, cell_address: str, new_value: Any) -> bool:
        """
        Обновляет значение редактируемой ячейки.

        Args:
            sheet_id (int): ID листа в БД.
            sheet_name (str): Имя листа Excel.
            cell_address (str): Адрес ячейки (например, 'A1').
            new_value (Any): Новое значение ячейки.

        Returns:
            bool: True, если операция прошла успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return editable_data.update_editable_cell(conn, sheet_id, sheet_name, cell_address, new_value)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при обновлении редактируемой ячейки {cell_address} для листа '{sheet_name}' (ID: {sheet_id}): {e}", exc_info=True)
            return False

    # --- Методы для работы с формулами ---

    # Используют функции из src/storage/formulas.py

    def save_sheet_formulas(self, sheet_id: int, formulas_list: List[Dict[str, str]]) -> bool:
        """
        Сохраняет формулы листа в БД проекта.

        Args:
            sheet_id (int): ID листа в БД.
            formulas_list (List[Dict[str, str]]): Список словарей с 'cell_address' и 'formula'.

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return formulas.save_sheet_formulas(conn, sheet_id, formulas_list)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при сохранении формул для листа ID {sheet_id}: {e}", exc_info=True)
            return False

    def load_sheet_formulas(self, sheet_id: int) -> List[Dict[str, str]]:
        """
        Загружает формулы листа из БД проекта.

        Args:
            sheet_id (int): ID листа в БД.

        Returns:
            List[Dict[str, str]]: Список словарей с 'cell_address' и 'formula'.
            Возвращает пустой список в случае ошибки или отсутствия данных.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    return formulas.load_sheet_formulas(conn, sheet_id)
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке формул для листа ID {sheet_id}: {e}", exc_info=True)
            return []

    # --- Методы для работы со стилями ---

    # Используют функции из src/storage/styles.py

    def save_sheet_styles(self, sheet_id: int, styles_list: List[Dict[str, Any]]) -> bool:
        """
        Сохраняет стили листа в БД проекта.

        Args:
            sheet_id (int): ID листа в БД.
            styles_list (List[Dict[str, Any]]): Список словарей с 'style_attributes' и 'range_address'.

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # Предполагается, что функция в styles.py имеет эту сигнатуру
                    return styles.save_sheet_styles(conn, sheet_id, styles_list)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при сохранении стилей для листа ID {sheet_id}: {e}", exc_info=True)
            return False

    def load_sheet_styles(self, sheet_id: int) -> List[Dict[str, Any]]:
        """
        Загружает стили и диапазоны для указанного листа.

        Args:
            sheet_id (int): ID листа в БД.

        Returns:
            List[Dict[str, Any]]: Список словарей с 'style_attributes' и 'range_address'.
            Возвращает пустой список в случае ошибки или отсутствия данных.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # Предполагается, что функция в styles.py имеет эту сигнатуру
                    return styles.load_sheet_styles(conn, sheet_id)
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке стилей для листа ID {sheet_id}: {e}", exc_info=True)
            return []

    # --- Методы для работы с диаграммами ---

    # Используют функции из src/storage/charts.py

    def save_sheet_charts(self, sheet_id: int, charts_list: List[Dict[str, Any]]) -> bool:
        """
        Сохраняет диаграммы листа в БД проекта.

        Args:
            sheet_id (int): ID листа в БД.
            charts_list (List[Dict[str, Any]]): Список словарей с данными диаграмм.

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # Предполагается, что функция в charts.py имеет эту сигнатуру
                    return charts.save_sheet_charts(conn, sheet_id, charts_list)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при сохранении диаграмм для листа ID {sheet_id}: {e}", exc_info=True)
            return False

    def load_sheet_charts(self, sheet_id: int) -> List[Dict[str, Any]]:
        """
        Загружает диаграммы для указанного листа.

        Args:
            sheet_id (int): ID листа в БД.

        Returns:
            List[Dict[str, Any]]: Список словарей с данными диаграмм.
            Возвращает пустой список в случае ошибки или отсутствия данных.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # Предполагается, что функция в charts.py имеет эту сигнатуру
                    return charts.load_sheet_charts(conn, sheet_id)
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке диаграмм для листа ID {sheet_id}: {e}", exc_info=True)
            return []

    # --- Методы для работы с историей редактирования ---

    # Используют функции из src/storage/history.py

    def save_edit_history_record(self, sheet_id: int, cell_address: str, old_value: Any, new_value: Any) -> bool:
        """
        Сохраняет запись в истории редактирования.

        Args:
            sheet_id (int): ID листа в БД.
            cell_address (str): Адрес ячейки.
            old_value (Any): Старое значение.
            new_value (Any): Новое значение.

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # Предполагается, что функция в history.py имеет эту сигнатуру
                    return history.save_edit_history_record(conn, sheet_id, cell_address, old_value, new_value)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при сохранении записи истории для листа ID {sheet_id}, ячейка {cell_address}: {e}", exc_info=True)
            return False

    def load_edit_history(self, sheet_id: Optional[int] = None, limit: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        Загружает историю редактирования.

        Args:
            sheet_id (Optional[int]): ID листа для фильтрации. Если None, загружает всю историю.
            limit (Optional[int]): Максимальное количество записей для загрузки.

        Returns:
            List[Dict[str, Any]]: Список записей истории.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # Предполагается, что функция в history.py имеет эту сигнатуру
                    return history.load_edit_history(conn, sheet_id, limit)
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке истории редактирования: {e}", exc_info=True)
            return []

    # --- Методы для работы с объединенными ячейками ---

    def load_sheet_merged_cells(self, sheet_id: int) -> List[str]:
        """
        Загружает список объединенных ячеек для листа.

        Args:
            sheet_id (int): ID листа в БД.

        Returns:
            List[str]: Список строковых адресов объединенных диапазонов (например, ['A1:B2', 'C3:D5']).
                    Возвращает пустой список, если данных нет или произошла ошибка.
        """
        if not self.connection:
            logger.error("[ОБЪЕДИНЕНИЕ] Нет подключения к БД для загрузки объединенных ячеек.")
            return []

        try:
            logger.debug(f"[ОБЪЕДИНЕНИЕ] Запрос объединенных ячеек для sheet_id={sheet_id}...")
            cursor = self.connection.execute(
                "SELECT merged_cells_data FROM sheet_merged_cells WHERE sheet_id = ?", (sheet_id,)
            )
            row = cursor.fetchone()
            
            if row and row['merged_cells_data']:
                merged_cells_list = json.loads(row['merged_cells_data'])
                if isinstance(merged_cells_list, list):
                    logger.info(f"[ОБЪЕДИНЕНИЕ] Загружено {len(merged_cells_list)} объединенных диапазонов для sheet_id={sheet_id}.")
                    return merged_cells_list
                else:
                    logger.warning(f"[ОБЪЕДИНЕНИЕ] Данные merged_cells_data для sheet_id={sheet_id} не являются списком.")
            else:
                logger.info(f"[ОБЪЕДИНЕНИЕ] Объединенные ячейки для sheet_id={sheet_id} не найдены.")

        except json.JSONDecodeError as je:
            logger.error(f"[ОБЪЕДИНЕНИЕ] Ошибка разбора JSON merged_cells_data для sheet_id={sheet_id}: {je}")
        except sqlite3.Error as e:
            logger.error(f"[ОБЪЕДИНЕНИЕ] Ошибка SQLite при загрузке merged_cells для sheet_id={sheet_id}: {e}")
        except Exception as e:
            logger.error(f"[ОБЪЕДИНЕНИЕ] Неожиданная ошибка при загрузке merged_cells для sheet_id={sheet_id}: {e}", exc_info=True)

        return []

    # Дополнительные методы и логика класса могут быть добавлены здесь