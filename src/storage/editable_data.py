# src/storage/editable_data.py

"""
Модуль для работы с редактируемыми данными в хранилище проекта Excel Micro DB.
"""

import sqlite3
import logging
from typing import Optional, Any, Dict, List
# Импортируем datetime для проверки типов значений
from datetime import datetime as dt_datetime # Переименовываем во избежание конфликта с именем переменной
# from src.storage.base import sanitize_table_name # Если потребуется
# from src.storage.schema import ... # Если потребуются какие-либо константы

logger = logging.getLogger(__name__)

# --- Вспомогательные функции ---

def sanitize_editable_table_name(name: str) -> str:
    """
    Санитизирует имя для использования в качестве имени таблицы для редактируемых данных.

    Args:
        name (str): Исходное имя.

    Returns:
        str: Санитизированное имя.
    """
    # Пока используем ту же функцию. Можно адаптировать при необходимости.
    # Импортируем здесь, чтобы избежать циклического импорта
    from src.storage.base import sanitize_table_name
    sanitized = sanitize_table_name(name)
    return f"editable_data_{sanitized}"

# --- Основные функции работы с редактируемыми данными ---

def load_sheet_editable_data(connection: sqlite3.Connection, sheet_name: str) -> Dict[str, Any]:
    """
    Загружает редактируемые данные для конкретного листа из его таблицы.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_name (str): Имя листа.

    Returns:
        Dict[str, Any]: Словарь с ключами 'column_names' и 'rows'.
                        Возвращает пустые списки, если таблица не найдена.
    """
    if not connection:
        logger.error("Получено пустое соединение для загрузки редактируемых данных.")
        return {"column_names": [], "rows": []}

    editable_table_name = sanitize_editable_table_name(sheet_name)
    
    try:
        cursor = connection.cursor()
        
        # Проверим, существует ли таблица
        cursor.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name=?
        """, (editable_table_name,))
        
        if not cursor.fetchone():
            logger.warning(f"Таблица редактируемых данных '{editable_table_name}' не найдена.")
            return {"column_names": [], "rows": []}

        # Получаем имена столбцов
        cursor.execute(f"PRAGMA table_info({editable_table_name})")
        columns_info = cursor.fetchall()
        # Пропускаем столбец 'id', который является служебным
        column_names = [col[1] for col in columns_info if col[1] != 'id']
        
        # Получаем все строки
        # Экранируем имена столбцов на случай, если они зарезервированы
        if column_names:
            placeholders = ', '.join([f'"{col}"' for col in column_names])
            cursor.execute(f"SELECT {placeholders} FROM {editable_table_name}")
            rows = cursor.fetchall()
        else:
            rows = []
        
        logger.debug(f"Загружено {len(rows)} строк из таблицы '{editable_table_name}'.")
        return {
            "column_names": column_names,
            "rows": rows
        }
        
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке редактируемых данных для листа '{sheet_name}': {e}")
        return {"column_names": [], "rows": []}
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке редактируемых данных для листа '{sheet_name}': {e}")
        return {"column_names": [], "rows": []}


def update_editable_cell(
    connection: sqlite3.Connection, 
    sheet_name: str, 
    row_index: int, 
    column_name: str, 
    new_value: Any
) -> bool:
    """
    Обновляет значение отдельной ячейки в таблице редактируемых данных конкретного листа.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_name (str): Имя листа.
        row_index (int): 0-базовый индекс строки.
        column_name (str): Имя столбца.
        new_value (Any): Новое значение ячейки.

    Returns:
        bool: True, если обновление прошло успешно, False в случае ошибки.
    """
    if not connection:
        logger.error("Получено пустое соединение для обновления редактируемой ячейки.")
        return False

    try:
        cursor = connection.cursor()
        
        # Санитизируем имя таблицы
        editable_table_name = sanitize_editable_table_name(sheet_name)
        
        # Санитизируем имя столбца
        # ИМПОРТ ВНУТРИ ФУНКЦИИ
        from src.storage.base import sanitize_column_name
        sanitized_col_name = sanitize_column_name(column_name)
        # Защита от конфликта с зарезервированным словом 'id'
        if sanitized_col_name.lower() == 'id':
             sanitized_col_name = f"data_{sanitized_col_name}"
             
        # row_index в Python (0-based) -> id в БД (1-based)
        db_row_id = row_index + 1 

        # Обновляем значение в таблице
        # Экранируем имя столбца
        cursor.execute(f'''
            UPDATE {editable_table_name} 
            SET "{sanitized_col_name}" = ?
            WHERE id = ?
        ''', (str(new_value), db_row_id)) # Преобразуем значение в строку для единообразия
        
        if cursor.rowcount == 0:
            logger.warning(f"Ячейка [{sheet_name}][{row_index}, {column_name}] не найдена для обновления (id={db_row_id} в таблице {editable_table_name}).")
            # Можно здесь создать новую строку, если она отсутствует, но обычно она должна быть.
            # Для MVP предполагаем, что строка существует.
            connection.commit() # На всякий случай коммитим, даже если ничего не изменилось
            return False
        else:
            connection.commit()
            logger.info(f"Обновлена ячейка [{sheet_name}][{row_index}, {column_name}] (id={db_row_id}). Новое значение: {new_value}")
            return True
            
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при обновлении ячейки [{sheet_name}][{row_index}, {column_name}]: {e}")
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при обновлении ячейки [{sheet_name}][{row_index}, {column_name}]: {e}")
        connection.rollback()
        return False

# === НОВОЕ: Функция для создания и инициализации таблицы редактируемых данных ===

def create_and_populate_editable_table(
    connection: sqlite3.Connection, 
    sheet_id: int, 
    sheet_name: str, 
    raw_data_info: Dict[str, Any]
) -> bool:
    """
    Создает таблицу editable_data_<sheet_name> и копирует в неё данные из raw_data_info.
    Если таблица уже существует, она не пересоздается.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа в БД.
        sheet_name (str): Имя листа Excel.
        raw_data_info (Dict[str, Any]): Словарь с ключами 'column_names' и 'rows', 
                                        содержащий сырые данные для копирования.

    Returns:
        bool: True, если таблица создана/заполнена успешно или уже существовала, False в случае ошибки.
    """
    # Инициализируем переменную заранее, чтобы Pylance был доволен
    editable_table_name = ""
    if not connection:
        logger.error("Получено пустое соединение для создания таблицы редактируемых данных.")
        return False

    try:
        cursor = connection.cursor()
        editable_table_name = sanitize_editable_table_name(sheet_name) # Определяем имя таблицы
        
        # 1. Проверить, существует ли таблица
        cursor.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name=?
        """, (editable_table_name,))
        
        if cursor.fetchone():
            logger.debug(f"Таблица редактируемых данных '{editable_table_name}' уже существует.")
            return True # Таблица уже есть, ничего не делаем
        
        logger.debug(f"Создание таблицы редактируемых данных '{editable_table_name}' для листа '{sheet_name}'.")

        # 2. Создать таблицу на основе raw_data_info
        column_names = raw_data_info.get("column_names", [])
        if not column_names:
             logger.warning(f"Нет имен столбцов в raw_data_info для листа '{sheet_name}'. Таблица не будет создана.")
             return True # Нечего создавать, но это не ошибка

        # Начинаем с обязательных служебных столбцов
        columns_sql_parts = ["id INTEGER PRIMARY KEY AUTOINCREMENT"] # Уникальный ID строки

        # Добавляем столбцы для данных, используя ту же логику санитизации
        for col_name in column_names:
            # Санитизируем имя столбца
            # ИМПОРТ ВНУТРИ ЦИКЛА
            from src.storage.base import sanitize_column_name
            sanitized_col_name = sanitize_column_name(col_name)
            # - ИСПРАВЛЕНО: Проверка на конфликт имён -
            # Проверяем, не совпадает ли санитизированное имя с зарезервированными
            if sanitized_col_name.lower() in ['id']:
                # Если совпадает, добавляем префикс
                sanitized_col_name = f"data_{sanitized_col_name}"
                logger.debug(f"Зарезервированное имя столбца '{col_name}' переименовано в '{sanitized_col_name}' для таблицы '{editable_table_name}'.")
            # - КОНЕЦ ИСПРАВЛЕНИЯ -
            # Добавляем столбец (в SQLite все значения TEXT, можно уточнить тип позже)
            columns_sql_parts.append(f"{sanitized_col_name} TEXT")

        create_table_sql = f"CREATE TABLE {editable_table_name} ({', '.join(columns_sql_parts)})"
        logger.debug(f"SQL-запрос создания таблицы: {create_table_sql}")
        cursor.execute(create_table_sql)
        
        # 3. Скопировать данные из raw_data_info
        logger.debug(f"Начало копирования данных в '{editable_table_name}' из переданного raw_data_info.")
        rows_data = raw_data_info.get("rows", [])
        if not rows_data:
             logger.debug(f"Нет строк данных в raw_data_info для листа '{sheet_name}'. Таблица создана пустая.")
             connection.commit()
             return True # Таблица создана, данных нет
        
        # Подготавливаем имена столбцов для вставки (санитизированные)
        sanitized_col_names = []
        for cn in column_names:
            # ИМПОРТ ВНУТРИ ЦИКЛА
            from src.storage.base import sanitize_column_name
            s_cn = sanitize_column_name(cn)
            # Применяем ту же логику переименования, что и при создании
            if s_cn.lower() == 'id': 
                s_cn = f"data_{s_cn}"
            sanitized_col_names.append(s_cn)

        logger.debug(f"Санитизированные имена столбцов для вставки: {sanitized_col_names}")

        if not sanitized_col_names:
            logger.warning(f"Нет санитизированных имен столбцов для вставки в '{editable_table_name}'.")
            connection.commit()
            return True # Таблица создана, но не заполнена
        
        # Формируем SQL-запрос для вставки
        placeholders = ', '.join(['?' for _ in sanitized_col_names])
        columns_str = ', '.join(sanitized_col_names)
        insert_sql = f"INSERT INTO {editable_table_name} ({columns_str}) VALUES ({placeholders})"
        logger.debug(f"SQL-запрос вставки данных: {insert_sql}")

        # Подготавливаем данные для вставки
        data_to_insert = []
        for row_dict in rows_data:
            # Для каждой строки создаем кортеж значений в порядке sanitized_col_names
            row_values = []
            for orig_col_name in column_names: # Итерируемся по оригинальным именам
                # ИМПОРТ ВНУТРИ ЦИКЛА
                from src.storage.base import sanitize_column_name
                sanitized_name = sanitize_column_name(orig_col_name)
                if sanitized_name.lower() == 'id':
                    sanitized_name = f"data_{sanitized_name}"
                # Получаем значение из row_dict по оригинальному имени
                value = row_dict.get(orig_col_name, None)
                # Убедимся, что значение является строкой или None для SQLite
                if value is not None and not isinstance(value, (str, int, float, type(None))):
                    # Если это datetime, преобразуем в ISO строку
                    # ИСПРАВЛЕНО: Импорт datetime и проверка типа
                    if isinstance(value, dt_datetime): # <-- ИСПРАВЛЕНО
                        value = value.isoformat()
                    else:
                        # Для остальных типов - преобразуем в строку
                        value = str(value)
                row_values.append(value)
            data_to_insert.append(tuple(row_values))

        logger.debug(f"Подготовлено {len(data_to_insert)} строк для вставки.")

        # Выполняем массовую вставку
        if data_to_insert: # Проверяем, есть ли данные для вставки
            cursor.executemany(insert_sql, data_to_insert)
            
        connection.commit()
        logger.info(f"Таблица редактируемых данных '{editable_table_name}' успешно создана и заполнена {len(data_to_insert)} строками.")
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при создании/заполнении таблицы '{editable_table_name}': {e}") # <-- ИСПРАВЛЕНО: использование переменной
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при создании/заполнении таблицы '{editable_table_name}': {e}") # <-- ИСПРАВЛЕНО: использование переменной
        connection.rollback()
        return False

# =========================================================