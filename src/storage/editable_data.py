# src/storage/editable_data.py

"""
Модуль для работы с редактируемыми данными в хранилище проекта Excel Micro DB.
"""

import sqlite3
import logging
from typing import Optional, Any, Dict, List
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
