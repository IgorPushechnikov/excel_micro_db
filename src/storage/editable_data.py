# src/storage/editable_data.py

import sqlite3
import logging
from typing import List, Dict, Any, Optional

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

def _get_raw_data_table_name(sheet_name: str) -> str:
    """
    Генерирует имя таблицы для "сырых" данных листа.
    Args:
        sheet_name (str): Имя листа Excel.
    Returns:
        str: Имя таблицы в БД.
    """
    # Санитизация имени таблицы для безопасности
    sanitized_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '_')).rstrip()
    sanitized_sheet_name = sanitized_sheet_name.replace(' ', '_')
    return f"raw_data_{sanitized_sheet_name}"

def load_sheet_editable_data(connection: sqlite3.Connection, sheet_id: int, sheet_name: str) -> List[Dict[str, Any]]:
    """
    Загружает "сырые" (редактируемые) данные для указанного листа.
    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.
        sheet_name (str): Имя листа Excel.
    Returns:
        List[Dict[str, Any]]: Список словарей с ключами 'cell_address' и 'value'.
                              Возвращает пустой список в случае ошибки или отсутствия данных.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки редактируемых данных.")
        return []

    try:
        cursor = connection.cursor()
        table_name = _get_raw_data_table_name(sheet_name)

        # Проверяем, существует ли таблица
        cursor.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name=?
        """, (table_name,))
        if not cursor.fetchone():
            logger.info(f"Таблица 'сырых' данных '{table_name}' не найдена. Возвращается пустой список.")
            return []

        # Загружаем данные
        cursor.execute(f"SELECT cell_address, value FROM {table_name}")
        rows = cursor.fetchall()
        
        editable_data = [{"cell_address": row[0], "value": row[1]} for row in rows]
        logger.debug(f"Загружено {len(editable_data)} записей 'сырых' данных для листа '{sheet_name}' (ID: {sheet_id}).")
        return editable_data

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке 'сырых' данных для листа '{sheet_name}' (ID: {sheet_id}): {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке 'сырых' данных для листа '{sheet_name}' (ID: {sheet_id}): {e}", exc_info=True)
        return []

def update_editable_cell(connection: sqlite3.Connection, sheet_id: int, sheet_name: str, cell_address: str, new_value: Any) -> bool:
    """
    Обновляет значение редактируемой ячейки в таблице raw_data_<имя_листа>.
    Если запись для ячейки не существует, она создается.
    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.
        sheet_name (str): Имя листа Excel.
        cell_address (str): Адрес ячейки (например, 'A1').
        new_value (Any): Новое значение ячейки.
    Returns:
        bool: True, если операция прошла успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для обновления редактируемой ячейки.")
        return False

    try:
        cursor = connection.cursor()
        table_name = _get_raw_data_table_name(sheet_name)

        # Проверяем, существует ли таблица, и создаем её, если нет
        cursor.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name=?
        """, (table_name,))
        if not cursor.fetchone():
            logger.info(f"Таблица '{table_name}' не найдена, создаем новую.")
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {table_name} (
                    cell_address TEXT PRIMARY KEY,
                    value TEXT,
                    value_type TEXT
                )
            """)
            connection.commit()
            logger.debug(f"Таблица '{table_name}' создана.")

        # Вставляем или обновляем запись
        # Для совместимости с raw_data, добавляем заглушку для value_type
        cursor.execute(f"""
            INSERT OR REPLACE INTO {table_name} (cell_address, value, value_type)
            VALUES (?, ?, ?)
        """, (cell_address, str(new_value) if new_value is not None else None, 'str'))
        
        connection.commit()
        logger.debug(f"Обновлено значение ячейки {cell_address} в таблице '{table_name}' для листа '{sheet_name}' (ID: {sheet_id}). Новое значение: {new_value}")
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при обновлении ячейки {cell_address} для листа '{sheet_name}' (ID: {sheet_id}): {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при обновлении ячейки {cell_address} для листа '{sheet_name}' (ID: {sheet_id}): {e}", exc_info=True)
        return False

# Дополнительные функции для работы с редактируемыми данными (если потребуются) могут быть добавлены здесь