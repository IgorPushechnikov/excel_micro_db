# src/storage/formulas.py
"""
Модуль для работы с формулами в хранилище проекта Excel Micro DB.
"""
import sqlite3
import json
import logging
from typing import List, Dict, Any, Optional

# from src.storage.base import DateTimeEncoder # Если потребуется
# from src.storage.schema import ... # Если потребуются какие-либо константы

logger = logging.getLogger(__name__)

def save_formulas(connection: sqlite3.Connection, sheet_id: int, formulas_data: List[Dict[str, Any]]) -> bool:
    """
    Сохраняет формулы для листа.
    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа.
        formulas_data (List[Dict[str, Any]]): Список словарей с данными формул.
    Returns:
        bool: True, если формулы сохранены успешно, False в случае ошибки.
    """
    if not connection:
        logger.error("Получено пустое соединение для сохранения формул.")
        return False
    try:
        cursor = connection.cursor()
        
        # Сначала удаляем старые формулы для этого листа (операция UPSERT)
        cursor.execute('DELETE FROM formulas WHERE sheet_id = ?', (sheet_id,))
        logger.debug(f"Удалены старые формулы для листа ID {sheet_id}.")

        for formula_info in formulas_data:
            cell = formula_info.get("cell", "")
            formula = formula_info.get("formula", "")
            references = formula_info.get("references", [])
            # Сериализуем ссылки в JSON
            references_json = json.dumps(references, ensure_ascii=False) # Предполагается сериализация выше
            # === ИСПРАВЛЕНО: Экранирование имени столбца ===
            cursor.execute(
                'INSERT INTO formulas (sheet_id, cell, formula, "references") VALUES (?, ?, ?, ?)', # <-- "references" в кавычках
                (sheet_id, cell, formula, references_json)
            )
            # =================================================
        connection.commit()
        logger.info(f"Формулы для листа ID {sheet_id} сохранены успешно.")
        return True
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении формул для листа ID {sheet_id}: {e}")
        connection.rollback()
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении формул для листа ID {sheet_id}: {e}")
        connection.rollback()
        return False

def load_formulas(connection: sqlite3.Connection, sheet_id: int) -> List[Dict[str, Any]]:
    """
    Загружает формулы для указанного листа.
    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        sheet_id (int): ID листа.
    Returns:
        List[Dict[str, Any]]: Список словарей с данными формул.
    """
    formulas_data = []
    if not connection:
        logger.error("Получено пустое соединение для загрузки формул.")
        return formulas_data
    try:
        cursor = connection.cursor()
        # === ИСПРАВЛЕНО: Экранирование имени столбца ===
        cursor.execute('SELECT cell, formula, "references" FROM formulas WHERE sheet_id = ?', (sheet_id,)) # <-- "references" в кавычках
        # =================================================
        formulas_rows = cursor.fetchall()
        for row in formulas_rows:
            cell, formula, references_json = row
            references = []
            if references_json:
                try:
                    references = json.loads(references_json)
                except json.JSONDecodeError as e:
                    logger.error(f"Ошибка десериализации ссылок формулы в ячейке {cell}: {e}")
            formulas_data.append({
                'cell': cell,
                'formula': formula,
                'references': references
            })
        logger.info(f"Загружено {len(formulas_data)} формул для листа ID {sheet_id}.")
        return formulas_data
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке формул для листа ID {sheet_id}: {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке формул для листа ID {sheet_id}: {e}")
        return []
