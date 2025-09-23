# src/storage/formulas.py

import sqlite3
import logging
from typing import List, Dict, Any

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

# Имя общей таблицы для хранения формул всех листов проекта
FORMULAS_TABLE_NAME = "formulas"

def save_sheet_formulas(connection: sqlite3.Connection, sheet_id: int, formulas_list: List[Dict[str, str]]) -> bool:
    """
    Сохраняет формулы листа в БД проекта.
    Формулы хранятся в общей таблице 'formulas', связанной с листом по sheet_id.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.
        formulas_list (List[Dict[str, str]]): Список словарей с 'cell_address' и 'formula'.

    Returns:
        bool: True, если сохранение успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения формул.")
        return False

    if not isinstance(formulas_list, list):
        logger.error(f"Неверный тип данных для formulas_list. Ожидался list, получен {type(formulas_list)}.")
        return False

    try:
        cursor = connection.cursor()
        
        # Удаляем существующие формулы для этого листа, чтобы избежать дубликатов
        logger.debug(f"Удаление существующих формул для sheet_id {sheet_id}...")
        cursor.execute(f"DELETE FROM {FORMULAS_TABLE_NAME} WHERE sheet_id = ?", (sheet_id,))
        logger.debug(f"Удалено {cursor.rowcount} существующих записей формул для sheet_id {sheet_id}.")

        # Подготавливаем данные для вставки
        # Используем INSERT OR REPLACE для простоты и атомарности
        formulas_to_insert = [
            (sheet_id, item.get('cell_address'), item.get('formula'))
            for item in formulas_list
            if item.get('cell_address') and item.get('formula') # Пропускаем записи без адреса или формулы
        ]

        if formulas_to_insert:
            logger.debug(f"Подготовлено {len(formulas_to_insert)} формул для листа ID {sheet_id}.")
            cursor.executemany(
                f"INSERT OR REPLACE INTO {FORMULAS_TABLE_NAME} (sheet_id, cell_address, formula) VALUES (?, ?, ?)",
                formulas_to_insert
            )
            connection.commit()
            logger.info(f"Сохранено {len(formulas_to_insert)} формул для листа ID {sheet_id} в таблицу '{FORMULAS_TABLE_NAME}'.")
        else:
            logger.info(f"Нет формул для сохранения для листа ID {sheet_id}.")
            
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении формул для листа ID {sheet_id}: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении формул для листа ID {sheet_id}: {e}", exc_info=True)
        return False


def load_sheet_formulas(connection: sqlite3.Connection, sheet_id: int) -> List[Dict[str, str]]:
    """
    Загружает формулы для указанного листа из БД проекта.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.

    Returns:
        List[Dict[str, str]]: Список словарей с 'cell_address' и 'formula'.
                             Возвращает пустой список в случае ошибки или отсутствия данных.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки формул.")
        return []

    try:
        cursor = connection.cursor()
        
        # Загружаем формулы для этого листа из общей таблицы
        logger.debug(f"Загрузка формул для sheet_id {sheet_id} из таблицы '{FORMULAS_TABLE_NAME}'...")
        cursor.execute(
            f"SELECT cell_address, formula FROM {FORMULAS_TABLE_NAME} WHERE sheet_id = ?",
            (sheet_id,)
        )
        rows = cursor.fetchall()
        
        formulas_data = [
            {"cell_address": row[0], "formula": row[1]}
            for row in rows
        ]
        
        logger.debug(f"Загружено {len(formulas_data)} формул для листа ID {sheet_id}.")
        return formulas_data

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке формул для листа ID {sheet_id}: {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке формул для листа ID {sheet_id}: {e}", exc_info=True)
        return []

# Дополнительные функции для работы с формулами (если потребуются) могут быть добавлены здесь
