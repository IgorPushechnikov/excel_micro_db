# src/storage/history.py
"""
Модуль для работы с историей изменений в хранилище проекта Excel Micro DB.
"""
import sqlite3
import json
import logging
from typing import Any, Optional
from datetime import datetime

# from src.storage.base import DateTimeEncoder # Если потребуется
# from src.storage.schema import ... # Если потребуются какие-либо константы

logger = logging.getLogger(__name__)

# --- SQL для создания таблицы истории (если она еще не создана) ---
# Лучше держать это в schema.py, но если забыли, можно и здесь определить.
# CREATE_EDIT_HISTORY_TABLE = '''
#     CREATE TABLE IF NOT EXISTS edit_history (
#         id INTEGER PRIMARY KEY AUTOINCREMENT,
#         project_id INTEGER NOT NULL,
#         sheet_id INTEGER,
#         cell_address TEXT,
#         action_type TEXT NOT NULL, -- 'edit_cell', 'add_row', 'delete_row', 'apply_style' и т.д.
#         old_value TEXT, -- JSON или строка
#         new_value TEXT, -- JSON или строка
#         timestamp TEXT NOT NULL, -- ISO формат
#         user TEXT, -- Имя пользователя, если применимо
#         details TEXT, -- Дополнительная информация в JSON
#         FOREIGN KEY (project_id) REFERENCES projects (id),
#         FOREIGN KEY (sheet_id) REFERENCES sheets (id)
#     )
# '''
# TABLE_CREATION_QUERIES.append(CREATE_EDIT_HISTORY_TABLE) # Добавить в schema.py

# --- Основные функции работы с историей ---

def save_edit_history_record(
    connection: sqlite3.Connection,
    project_id: int,
    sheet_id: Optional[int],
    cell_address: Optional[str],
    action_type: str,
    old_value: Any,
    new_value: Any,
    user: Optional[str] = None,
    details: Optional[dict] = None
) -> bool:
    """
    Сохраняет запись об изменении в истории.
    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        project_id (int): ID проекта.
        sheet_id (Optional[int]): ID листа (если применимо).
        cell_address (Optional[str]): Адрес ячейки (если применимо).
        action_type (str): Тип действия (например, 'edit_cell').
        old_value (Any): Старое значение.
        new_value (Any): Новое значение.
        user (Optional[str]): Имя пользователя.
        details (Optional[dict]): Дополнительные детали.
    Returns:
        bool: True, если запись сохранена успешно, False в случае ошибки.
    """
    if not connection:
        logger.error("Получено пустое соединение для сохранения записи истории.")
        return False

    try:
        cursor = connection.cursor()
        
        # Преобразуем значения в строки/JSON, если необходимо
        # Для простоты преобразуем всё в строку. Можно использовать json.dumps с DateTimeEncoder.
        # serialized_old_value = json.dumps(old_value, cls=DateTimeEncoder, ensure_ascii=False) if old_value is not None else None
        # serialized_new_value = json.dumps(new_value, cls=DateTimeEncoder, ensure_ascii=False) if new_value is not None else None
        # serialized_details = json.dumps(details, cls=DateTimeEncoder, ensure_ascii=False) if details else None
        # Пока просто str(), предполагая, что это будет обработано выше или значения уже сериализованы.
        serialized_old_value = str(old_value) if old_value is not None else None
        serialized_new_value = str(new_value) if new_value is not None else None
        serialized_details = json.dumps(details, ensure_ascii=False) if details else None

        timestamp_iso = datetime.now().isoformat()

        cursor.execute('''
            INSERT INTO edit_history 
            (project_id, sheet_id, cell_address, action_type, old_value, new_value, timestamp, user, details)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            project_id, sheet_id, cell_address, action_type,
            serialized_old_value, serialized_new_value,
            timestamp_iso, user, serialized_details
        ))

        connection.commit()
        logger.debug(f"Запись истории сохранена: {action_type} на листе {sheet_id}, ячейка {cell_address}.")
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении записи истории: {e}")
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении записи истории: {e}")
        connection.rollback()
        return False
