# src/storage/editable_data.py
"""
Модуль для работы с редактируемыми данными в хранилище проекта Excel Micro DB.
(Пока содержит только функции для обновления отдельных ячеек)
"""
import sqlite3
import logging
from typing import Optional, Any

# from src.storage.base import sanitize_table_name # Если потребуется
# from src.storage.schema import ... # Если потребуются какие-либо константы

logger = logging.getLogger(__name__)

# --- Вспомогательные функции ---
def sanitize_editable_table_name(name: str) -> str:
    """
    Санитизирует имя для использования в качестве имени таблицы для редактируемых данных.
    Можно использовать ту же логику, что и для сырых данных, или свою.
    Args:
        name (str): Исходное имя.
    Returns:
        str: Санитизированное имя.
    """
    # Пока используем ту же функцию. Можно адаптировать при необходимости.
    from src.storage.base import sanitize_table_name
    sanitized = sanitize_table_name(name)
    return f"editable_{sanitized}"

# --- Основные функции работы с редактируемыми данными ---

# Предположим, что метод load_sheet_editable_data будет добавлен позже.
# def load_sheet_editable_data(...): pass

def update_editable_cell(connection: sqlite3.Connection, sheet_id: int, cell_address: str, new_value: Any) -> bool:
    """
    Обновляет значение отдельной ячейки в таблице редактируемых данных.
    (Предполагается, что таблица уже существует и имеет структуру cell_address -> value)
    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа.
        cell_address (str): Адрес ячейки (например, 'A1').
        new_value (Any): Новое значение ячейки.
    Returns:
        bool: True, если обновление прошло успешно, False в случае ошибки.
    """
    # Эта функция требует доработки, так как структура таблицы редактируемых данных
    # не определена в исходном database.py. Это просто пример.
    # Предположим, что есть таблица editable_data с полями:
    # sheet_id, cell_address, value, last_modified

    if not connection:
        logger.error("Получено пустое соединение для обновления редактируемой ячейки.")
        return False

    try:
        cursor = connection.cursor()
        # Простая логика UPSERT (в SQLite 3.24.0+ можно использовать INSERT ... ON CONFLICT)
        # Используем INSERT OR REPLACE для простоты
        cursor.execute('''
            INSERT OR REPLACE INTO editable_data (sheet_id, cell_address, value, last_modified)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ''', (sheet_id, cell_address, str(new_value))) # Преобразуем значение в строку
        
        connection.commit()
        logger.info(f"Обновлена ячейка {cell_address} на листе ID {sheet_id}. Новое значение: {new_value}")
        return True
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при обновлении ячейки {cell_address} на листе ID {sheet_id}: {e}")
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при обновлении ячейки {cell_address} на листе ID {sheet_id}: {e}")
        connection.rollback()
        return False
