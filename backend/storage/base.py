# backend/storage/base.py

import sqlite3
import logging
from contextlib import contextmanager
from typing import List, Dict, Any, Optional, Union
import os
import json

# Импортируем новые функции из модулей storage
# ИСПРАВЛЕНО: Все импорты теперь с префиксом backend.
from backend.storage import schema, raw_data, editable_data, formulas, styles, charts, history, metadata, sheets # <-- ИСПРАВЛЕНО: было from . import ...

# Импортируем logger из utils
# ИСПРАВЛЕНО: Импорт теперь из backend.utils
from backend.utils.logger import get_logger # <-- ИСПРАВЛЕНО: было from utils.logger

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
            # ИСПРАВЛЕНО: Вызов schema.initialize_project_schema теперь с префиксом backend.storage
            schema.initialize_project_schema(self.connection) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов sheets.save_sheet теперь с префиксом backend.storage
                    return sheets.save_sheet(conn, project_id, sheet_name, max_row, max_column) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов sheets.load_all_sheets_metadata теперь с префиксом backend.storage
                    return sheets.load_all_sheets_metadata(conn, project_id) # <-- ИСПРАВЛЕНО
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке списка листов для проекта ID {project_id}: {e}", exc_info=True)
            return []

    # --- ОБНОВЛЕНО: Метод для переименования листа ---
    def rename_sheet(self, project_id: int, old_name: str, new_name: str) -> bool:
        """
        Переименовывает лист в таблице 'sheets' и обновляет связанные таблицы.

        Args:
            project_id (int): ID проекта.
            old_name (str): Текущее имя листа.
            new_name (str): Новое имя листа.

        Returns:
            bool: True, если переименование успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # Вызов sheets.rename_sheet
                    return sheets.rename_sheet(conn, project_id, old_name, new_name)
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при переименовании листа '{old_name}' в '{new_name}': {e}", exc_info=True)
            return False
    # --- КОНЕЦ ОБНОВЛЕНИЯ ---

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
                    # ИСПРАВЛЕНО: Вызов metadata.save_sheet_metadata теперь с префиксом backend.storage
                    return metadata.save_sheet_metadata(conn, sheet_name, sheet_data) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов metadata.load_sheet_metadata теперь с префиксом backend.storage
                    return metadata.load_sheet_metadata(conn, sheet_name) # <-- ИСПРАВЛЕНО
                else:
                    return None
        except Exception as e:
            logger.error(f"Ошибка при загрузке метаданных листа '{sheet_name}': {e}", exc_info=True)
            return None

    # --- НОВОЕ: Метод для сохранения метаданных проекта ---
    def save_project_metadata(self, project_id: int, metadata_dict: Dict[str, Any]) -> bool:
        """
        Сохраняет метаданные проекта в БД.

        Args:
            project_id (int): ID проекта.
            metadata_dict (Dict[str, Any]): Словарь с метаданными проекта.

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            with self.get_connection() as conn:
                if conn:
                    # ИСПРАВЛЕНО: Вызов metadata.save_project_metadata теперь с префиксом backend.storage
                    return metadata.save_project_metadata(conn, project_id, metadata_dict) # <-- ИСПРАВЛЕНО
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при сохранении метаданных проекта (ID: {project_id}): {e}", exc_info=True)
            return False
    # --- КОНЕЦ НОВОГО ---

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
                    # ИСПРАВЛЕНО: Вызов raw_data.save_sheet_raw_data теперь с префиксом backend.storage
                    return raw_data.save_sheet_raw_data(conn, sheet_name, raw_data_list) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов raw_data.load_sheet_raw_data теперь с префиксом backend.storage
                    return raw_data.load_sheet_raw_data(conn, sheet_name) # <-- ИСПРАВЛЕНО
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке сырых данных листа '{sheet_name}': {e}", exc_info=True)
            return []

    # --- Методы для работы с редактируемыми данными ---

    # Используют функции из storage/editable_data.py

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
                    # ИСПРАВЛЕНО: Вызов editable_data.load_sheet_editable_data теперь с префиксом backend.storage
                    return editable_data.load_sheet_editable_data(conn, sheet_id, sheet_name) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов editable_data.update_editable_cell теперь с префиксом backend.storage
                    return editable_data.update_editable_cell(conn, sheet_id, sheet_name, cell_address, new_value) # <-- ИСПРАВЛЕНО
                else:
                    return False
        except Exception as e:
            logger.error(f"Ошибка при обновлении редактируемой ячейки {cell_address} для листа '{sheet_name}' (ID: {sheet_id}): {e}", exc_info=True)
            return False

    # --- Методы для работы с формулами ---

    # Используют функции из storage/formulas.py

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
                    # ИСПРАВЛЕНО: Вызов formulas.save_sheet_formulas теперь с префиксом backend.storage
                    return formulas.save_sheet_formulas(conn, sheet_id, formulas_list) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов formulas.load_sheet_formulas теперь с префиксом backend.storage
                    return formulas.load_sheet_formulas(conn, sheet_id) # <-- ИСПРАВЛЕНО
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке формул для листа ID {sheet_id}: {e}", exc_info=True)
            return []

    # --- Методы для работы со стилями ---

    # Используют функции из storage/styles.py

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
                    # ИСПРАВЛЕНО: Вызов styles.save_sheet_styles теперь с префиксом backend.storage
                    return styles.save_sheet_styles(conn, sheet_id, styles_list) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов styles.load_sheet_styles теперь с префиксом backend.storage
                    return styles.load_sheet_styles(conn, sheet_id) # <-- ИСПРАВЛЕНО
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке стилей для листа ID {sheet_id}: {e}", exc_info=True)
            return []

    # --- Методы для работы с диаграммами ---

    # Используют функции из storage/charts.py

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
                    # ИСПРАВЛЕНО: Вызов charts.save_sheet_charts теперь с префиксом backend.storage
                    return charts.save_sheet_charts(conn, sheet_id, charts_list) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов charts.load_sheet_charts теперь с префиксом backend.storage
                    return charts.load_sheet_charts(conn, sheet_id) # <-- ИСПРАВЛЕНО
                else:
                    return []
        except Exception as e:
            logger.error(f"Ошибка при загрузке диаграмм для листа ID {sheet_id}: {e}", exc_info=True)
            return []

    # --- Методы для работы с объединенными ячейками ---

    def save_sheet_merged_cells(self, sheet_id: int, merged_cells_list: List[str]) -> bool:
        """
        Сохраняет список объединенных ячеек для листа в БД проекта.

        Args:
            sheet_id (int): ID листа в БД.
            merged_cells_list (List[str]): Список строковых адресов диапазонов (например, ['A1:B2', 'C3:D5']).

        Returns:
            bool: True, если сохранение успешно, иначе False.
        """
        try:
            # Импортируем json здесь, если еще не импортирован глобально в этом файле
            # import json 
            
            with self.get_connection() as conn:
                if conn:
                    serialized_data = json.dumps(merged_cells_list)
                    cursor = conn.cursor()
                    
                    # Используем INSERT OR REPLACE для обновления или вставки
                    cursor.execute(
                        """
                        INSERT OR REPLACE INTO sheet_merged_cells (sheet_id, merged_cells_data)
                        VALUES (?, ?)
                        """,
                        (sheet_id, serialized_data)
                    )
                    conn.commit()
                    logger.info(f"[ОБЪЕДИНЕНИЕ] Сохранено {len(merged_cells_list)} объединенных диапазонов для sheet_id={sheet_id}.")
                    return True
                else:
                    return False
        # ИСПРАВЛЕНО: Заменено json.JSONEncodeError на ValueError и TypeError
        except ValueError as je: # json.dumps может выбросить ValueError для недопустимых типов
            logger.error(f"[ОБЪЕДИНЕНИЕ] Ошибка сериализации (ValueError) списка объединенных ячеек для sheet_id={sheet_id}: {je}")
        except TypeError as te: # или TypeError
            logger.error(f"[ОБЪЕДИНЕНИЕ] Ошибка сериализации (TypeError) списка объединенных ячеек для sheet_id={sheet_id}: {te}")
        # ИСПРАВЛЕНО: Заменено json.JSONEncodeError на ValueError
        except sqlite3.Error as e:
            logger.error(f"[ОБЪЕДИНЕНИЕ] Ошибка SQLite при сохранении merged_cells для sheet_id={sheet_id}: {e}")
        except Exception as e:
            logger.error(f"[ОБЪЕДИНЕНИЕ] Неожиданная ошибка при сохранении merged_cells для sheet_id={sheet_id}: {e}", exc_info=True)
        
        return False

    # --- Методы для работы с историей редактирования ---

    # Используют функции из storage/history.py

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
                    # ИСПРАВЛЕНО: Вызов history.save_edit_history_record теперь с префиксом backend.storage
                    return history.save_edit_history_record(conn, sheet_id, cell_address, old_value, new_value) # <-- ИСПРАВЛЕНО
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
                    # ИСПРАВЛЕНО: Вызов history.load_edit_history теперь с префиксом backend.storage
                    return history.load_edit_history(conn, sheet_id, limit) # <-- ИСПРАВЛЕНО
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
            # Устанавливаем row_factory для получения sqlite3.Row
            self.connection.row_factory = sqlite3.Row
            cursor = self.connection.execute(
                "SELECT merged_cells_data FROM sheet_merged_cells WHERE sheet_id = ?", (sheet_id,)
            )
            row = cursor.fetchone()
            # Сбрасываем row_factory в значение по умолчанию (None -> tuple)
            self.connection.row_factory = None
            
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
