# src/storage/history.py

import sqlite3
import logging
from typing import List, Dict, Any, Optional

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

def save_edit_history_record(
    connection: sqlite3.Connection, 
    sheet_id: int, 
    cell_address: str, 
    old_value: Any, 
    new_value: Any
) -> bool:
    """
    Сохраняет запись об изменении ячейки в истории редактирования.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа, где произошло изменение.
        cell_address (str): Адрес ячейки (например, 'A1').
        old_value (Any): Предыдущее значение ячейки.
        new_value (Any): Новое значение ячейки.

    Returns:
        bool: True, если запись успешно сохранена, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения записи истории.")
        return False

    try:
        cursor = connection.cursor()
        
        # Получаем project_id из sheet_id для полноты записи
        cursor.execute("SELECT project_id FROM sheets WHERE sheet_id = ?", (sheet_id,))
        result = cursor.fetchone()
        if not result:
            logger.error(f"Не найден project_id для sheet_id {sheet_id}. Запись истории не сохранена.")
            return False
            
        project_id = result[0]

        # Вставляем запись в таблицу edit_history
        # Предполагается, что таблица edit_history имеет поля:
        # history_id (INTEGER PRIMARY KEY AUTOINCREMENT)
        # project_id (INTEGER)
        # sheet_id (INTEGER)
        # cell_address (TEXT)
        # old_value (TEXT) -- Можно хранить как строку
        # new_value (TEXT) -- Можно хранить как строку
        # edited_at (TEXT DEFAULT (datetime('now'))) -- Время автоматически
        
        cursor.execute("""
            INSERT INTO edit_history (project_id, sheet_id, cell_address, old_value, new_value)
            VALUES (?, ?, ?, ?, ?)
        """, (project_id, sheet_id, cell_address, str(old_value), str(new_value)))
        
        connection.commit()
        logger.debug(f"Запись истории редактирования сохранена для ячейки {cell_address} на листе ID {sheet_id}.")
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении записи истории для ячейки {cell_address} на листе ID {sheet_id}: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении записи истории для ячейки {cell_address} на листе ID {sheet_id}: {e}", exc_info=True)
        return False


def load_edit_history(
    connection: sqlite3.Connection, 
    sheet_id: Optional[int] = None, 
    limit: Optional[int] = None
) -> List[Dict[str, Any]]:
    """
    Загружает историю редактирования.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (Optional[int]): ID листа для фильтрации. 
                                  Если None, загружает всю историю проекта.
        limit (Optional[int]): Максимальное количество записей для загрузки.

    Returns:
        List[Dict[str, Any]]: Список словарей, представляющих записи истории.
                             Каждый словарь содержит ключи: history_id, project_id, sheet_id, 
                             cell_address, old_value, new_value, edited_at.
                             Возвращает пустой список в случае ошибки.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки истории редактирования.")
        return []

    try:
        cursor = connection.cursor()
        
        # Формируем SQL-запрос с возможной фильтрацией и ограничением
        query = "SELECT history_id, project_id, sheet_id, cell_address, old_value, new_value, edited_at FROM edit_history"
        params = []
        
        if sheet_id is not None:
            query += " WHERE sheet_id = ?"
            params.append(sheet_id)
            
        query += " ORDER BY edited_at DESC, history_id DESC" # Новые записи первыми
        
        if limit is not None:
            query += " LIMIT ?"
            params.append(limit)
            
        cursor.execute(query, params)
        rows = cursor.fetchall()
        
        history_records = []
        for row in rows:
            history_records.append({
                "history_id": row[0],
                "project_id": row[1],
                "sheet_id": row[2],
                "cell_address": row[3],
                "old_value": row[4],
                "new_value": row[5],
                "edited_at": row[6]
            })
            
        logger.debug(f"Загружено {len(history_records)} записей истории редактирования (sheet_id={sheet_id}, limit={limit}).")
        return history_records

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке истории редактирования (sheet_id={sheet_id}, limit={limit}): {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке истории редактирования (sheet_id={sheet_id}, limit={limit}): {e}", exc_info=True)
        return []

# Дополнительные функции для работы с историей (например, откат по ID) могут быть добавлены здесь
