# src/storage/raw_data.py

"""
Модуль для работы с сырыми данными в хранилище проекта Excel Micro DB.

Содержит логику создания таблиц, сохранения и загрузки сырых данных.
"""

import sqlite3
import logging
from typing import List, Dict, Any, Optional
from datetime import datetime
# from src.storage.base import sanitize_table_name # ИМПОРТ ПЕРЕМЕЩЕН ВНУТРЬ ФУНКЦИЙ, чтобы избежать циклического импорта

# from src.storage.schema import ... # Если потребуются какие-либо константы из schema, импортируем их

logger = logging.getLogger(__name__)

# --- Вспомогательные функции ---

def _get_raw_data_table_name(sheet_name: str) -> str:
    """Генерирует имя таблицы для сырых данных листа."""
    # Импортируем здесь, чтобы избежать циклического импорта
    from src.storage.base import sanitize_table_name
    # Используем санитизированное имя листа как основу
    base_name = sanitize_table_name(sheet_name)
    return f"raw_data_{base_name}"

# --- Основные функции работы с сырыми данными ---

def create_raw_data_table(connection: sqlite3.Connection, sheet_name: str, column_names: List[str]) -> bool:
    """
    Создает таблицу в БД для хранения сырых данных листа.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_name (str): Имя листа Excel.
        column_names (List[str]): Список имен столбцов.

    Returns:
        bool: True, если таблица создана или уже существует, False в случае ошибки.
    """
    # === ИСПРАВЛЕНО: Инициализация table_name в начале функции ===
    # Инициализируем table_name в начале области видимости функции, чтобы Pylance был доволен
    # и чтобы переменная была доступна в блоках except
    table_name = ""
    # === КОНЕЦ ИСПРАВЛЕНИЯ ===

    if not connection:
        logger.error("[DEBUG_STORAGE] Получено пустое соединение для создания таблицы сырых данных.")
        return False

    try:
        cursor = connection.cursor()
        # === ИСПРАВЛЕНО: Присваивание значения table_name внутри try ===
        # table_name будет определена здесь, если код дойдет до этой точки без ошибок
        table_name = _get_raw_data_table_name(sheet_name)
        # === КОНЕЦ ИСПРАВЛЕНИЯ ===
        logger.debug(f"[DEBUG_STORAGE] Создание таблицы сырых данных '{table_name}' для листа '{sheet_name}'.")

        # 1. Проверяем, существует ли таблица
        cursor.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name=?
        """, (table_name,))
        if cursor.fetchone():
            logger.debug(f"[DEBUG_STORAGE] Таблица '{table_name}' уже существует.")
            return True  # Таблица уже существует

        # 2. Создаем таблицу
        # Начинаем с обязательных служебных столбцов
        columns_sql_parts = ["id INTEGER PRIMARY KEY AUTOINCREMENT"]  # Уникальный ID строки

        # Добавляем столбцы для данных
        for col_name in column_names:
            # Санитизируем имя столбца
            # Импортируем здесь внутри цикла для избежания циклического импорта
            from src.storage.base import sanitize_table_name # Хотя это избыточно, но соответствует стилю файла
            sanitized_col_name = sanitize_table_name(col_name)
            # - ИСПРАВЛЕНО: Проверка на конфликт имён -
            # Проверяем, не совпадает ли санитизированное имя с зарезервированными
            if sanitized_col_name.lower() in ['id']:
                # Если совпадает, добавляем префикс
                sanitized_col_name = f"data_{sanitized_col_name}"
            logger.debug(f"[DEBUG_STORAGE] Зарезервированное имя столбца '{col_name}' переименовано в '{sanitized_col_name}' для таблицы '{table_name}'.")
            # - КОНЕЦ ИСПРАВЛЕНИЯ -
            # Добавляем столбец (в SQLite все значения TEXT, можно уточнить тип позже)
            columns_sql_parts.append(f"{sanitized_col_name} TEXT")

        create_table_sql = f"CREATE TABLE {table_name} ({', '.join(columns_sql_parts)})"
        logger.debug(f"[DEBUG_STORAGE] SQL-запрос создания таблицы: {create_table_sql}")
        cursor.execute(create_table_sql)

        # 3. Регистрируем таблицу в реестре
        # Регистрация будет происходить в save_sheet_raw_data.
        connection.commit()
        logger.info(f"[DEBUG_STORAGE] Таблица сырых данных '{table_name}' успешно создана.")
        return True

    except sqlite3.Error as e:
        # === ИСПРАВЛЕНО: table_name теперь всегда определена ===
        logger.error(f"[DEBUG_STORAGE] Ошибка SQLite при создании таблицы '{table_name}': {e}")
        # === КОНЕЦ ИСПРАВЛЕНИЯ ===
        connection.rollback()
        return False
    except Exception as e:
        # === ИСПРАВЛЕНО: table_name теперь всегда определена ===
        logger.error(f"[DEBUG_STORAGE] Неожиданная ошибка при создании таблицы '{table_name}': {e}")
        # === КОНЕЦ ИСПРАВЛЕНИЯ ===
        connection.rollback()
        return False

def save_sheet_raw_data(connection: sqlite3.Connection, sheet_id: int, sheet_name: str, raw_data_info: Dict[str, Any]) -> bool:
    """
    Сохраняет сырые данные листа в отдельную таблицу.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа в БД.
        sheet_name (str): Имя листа Excel.
        raw_data_info (Dict[str, Any]): Информация о сырых данных (ключи: 'column_names', 'rows').

    Returns:
        bool: True, если данные сохранены успешно, False в случае ошибки.
    """
    if not connection:
        logger.error("Получено пустое соединение для сохранения сырых данных.")
        return False

    try:
        table_name = _get_raw_data_table_name(sheet_name)
        column_names = raw_data_info.get("column_names", [])
        rows_data = raw_data_info.get("rows", [])

        logger.debug(f"[DEBUG_STORAGE] Начало сохранения сырых данных для листа '{sheet_name}' (ID: {sheet_id}) в таблицу '{table_name}'.")

        # 1. Создаем таблицу (если не существует)
        # Передаем column_names для создания столбцов
        if not create_raw_data_table(connection, sheet_name, column_names):
            logger.error(f"[DEBUG_STORAGE] Не удалось создать таблицу для сырых данных листа '{sheet_name}'.")
            return False

        # 2. Вставляем данные
        cursor = connection.cursor()

        # Подготавливаем имена столбцов для вставки (санитизированные)
        sanitized_col_names = []
        for cn in column_names:
            # Импортируем здесь внутри цикла для избежания циклического импорта
            from src.storage.base import sanitize_table_name # Хотя это избыточно, но соответствует стилю файла
            s_cn = sanitize_table_name(cn)
            # Применяем ту же логику переименования, что и при создании
            if s_cn.lower() == 'id':  # Только 'id' конфликтует, 'row_index' не создается как столбец
                s_cn = f"data_{s_cn}"
            sanitized_col_names.append(s_cn)

        logger.debug(f"[DEBUG_STORAGE] Санитизированные имена столбцов для вставки: {sanitized_col_names}")

        if not sanitized_col_names:
            logger.warning(f"[DEBUG_STORAGE] Нет санитизированных имен столбцов для вставки в '{table_name}'.")
            return True  # Нечего вставлять, но это не ошибка

        # Формируем SQL-запрос для вставки
        # placeholders - это список '?' для VALUES
        placeholders = ', '.join(['?' for _ in sanitized_col_names])
        # columns_str - это список санитизированных имен столбцов
        columns_str = ', '.join(sanitized_col_names)
        insert_sql = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        logger.debug(f"[DEBUG_STORAGE] SQL-запрос вставки данных: {insert_sql}")

        # Подготавливаем данные для вставки
        data_to_insert = []
        for row_dict in rows_data:
            # Для каждой строки создаем кортеж значений в порядке sanitized_col_names
            row_values = []
            for orig_col_name in column_names:  # Итерируемся по оригинальным именам
                # Импортируем здесь внутри цикла для избежания циклического импорта
                from src.storage.base import sanitize_table_name # Хотя это избыточно, но соответствует стилю файла
                sanitized_name = sanitize_table_name(orig_col_name)
                if sanitized_name.lower() == 'id':
                    sanitized_name = f"data_{sanitized_name}"
                # Получаем значение из row_dict по оригинальному имени
                value = row_dict.get(orig_col_name, None)
                # Убедимся, что значение является строкой или None для SQLite
                if value is not None and not isinstance(value, (str, int, float, type(None))):
                    # Если это datetime, преобразуем в ISO строку
                    if isinstance(value, datetime):
                        value = value.isoformat()
                    else:
                        # Для остальных типов - преобразуем в строку
                        value = str(value)
                row_values.append(value)
            data_to_insert.append(tuple(row_values))

        logger.debug(f"[DEBUG_STORAGE] Подготовлено {len(data_to_insert)} строк для вставки.")

        # Выполняем массовую вставку
        if data_to_insert:  # Проверяем, есть ли данные для вставки
            cursor.executemany(insert_sql, data_to_insert)

        # 3. Регистрируем таблицу в реестре (если ещё не зарегистрирована)
        cursor.execute("""
            INSERT OR IGNORE INTO raw_data_tables_registry (sheet_id, table_name)
            VALUES (?, ?)
        """, (sheet_id, table_name))
        connection.commit()

        logger.info(f"[DEBUG_STORAGE] В таблицу '{table_name}' вставлено {len(data_to_insert)} строк сырых данных и зарегистрирована.")

        logger.info(f"[DEBUG_STORAGE] Сырые данные для листа '{sheet_name}' (ID: {sheet_id}) успешно сохранены в таблицу '{table_name}'.")
        return True

    except sqlite3.Error as e:
        logger.error(f"[DEBUG_STORAGE] Ошибка SQLite при сохранении сырых данных для листа '{sheet_name}' (ID: {sheet_id}): {e}")
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"[DEBUG_STORAGE] Неожиданная ошибка при сохранении сырых данных для листа '{sheet_name}' (ID: {sheet_id}): {e}")
        connection.rollback()
        return False

def load_sheet_raw_data(connection: sqlite3.Connection, sheet_name: str) -> Dict[str, Any]:
    """
    Загружает сырые данные листа из его таблицы.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_name (str): Имя листа Excel.

    Returns:
        Dict[str, Any]: Словарь с ключами 'column_names' и 'rows'.
                        Возвращает пустые списки, если таблица не найдена.
    """
    raw_data_info = {"column_names": [], "rows": []}
    if not connection:
        logger.error("Получено пустое соединение для загрузки сырых данных.")
        return raw_data_info

    try:
        table_name = _get_raw_data_table_name(sheet_name)
        logger.debug(f"[DEBUG_STORAGE] Начало загрузки сырых данных для листа '{sheet_name}' из таблицы '{table_name}'.")

        # Проверяем, существует ли таблица
        cursor = connection.cursor()
        cursor.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name=?
        """, (table_name,))
        if not cursor.fetchone():
            logger.warning(f"[DEBUG_STORAGE] Таблица сырых данных '{table_name}' не найдена.")
            return raw_data_info  # Возвращаем пустую структуру

        # Получаем имена столбцов
        cursor.execute(f"PRAGMA table_info({table_name})")
        columns_info = cursor.fetchall()
        # columns_info - список кортежей (cid, name, type, notnull, dflt_value, pk)
        # Исключаем служебный столбец 'id'
        column_names = [col_info[1] for col_info in columns_info if col_info[1].lower() != 'id']
        # Убираем префикс 'data_' если он был добавлен
        original_column_names = [cn[5:] if cn.startswith('data_') else cn for cn in column_names]
        raw_data_info["column_names"] = original_column_names
        logger.debug(f"[DEBUG_STORAGE] Загружены имена столбцов: {original_column_names}")

        # Получаем все строки данных
        # Формируем список столбцов для SELECT
        select_columns = ', '.join(column_names) if column_names else '*'
        cursor.execute(f"SELECT {select_columns} FROM {table_name}")
        rows = cursor.fetchall()

        # Преобразуем кортежи в словари
        rows_data = []
        for row_tuple in rows:
            row_dict = {}
            for i, col_name in enumerate(column_names):
                orig_col_name = original_column_names[i]
                value = row_tuple[i]
                # Здесь можно добавить десериализацию из строки, если потребуется
                row_dict[orig_col_name] = value
            rows_data.append(row_dict)

        raw_data_info["rows"] = rows_data
        logger.info(f"[DEBUG_STORAGE] Сырые данные для листа '{sheet_name}' загружены. Всего строк: {len(raw_data_info['rows'])}")
        return raw_data_info

    except sqlite3.Error as e:
        logger.error(f"[DEBUG_STORAGE] Ошибка SQLite при загрузке сырых данных для листа '{sheet_name}': {e}")
        # Возвращаем пустую структуру в случае ошибки
        return {"column_names": [], "rows": []}
    except Exception as e:
        logger.error(f"[DEBUG_STORAGE] Неожиданная ошибка при загрузке сырых данных для листа '{sheet_name}': {e}")
        # Возвращаем пустую структуру в случае ошибки
        return {"column_names": [], "rows": []}
