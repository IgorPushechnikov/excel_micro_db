# src/storage/sheets.py

import sqlite3
import logging
from typing import Dict, Any, Optional, List

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)


def save_sheet(connection: sqlite3.Connection, project_id: int, sheet_name: str, max_row: Optional[int] = None, max_column: Optional[int] = None) -> Optional[int]:
    """
    Сохраняет информацию о листе в таблицу 'sheets'.
    Если лист с таким именем для проекта уже существует, возвращает его sheet_id.
    Иначе создает новую запись.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        project_id (int): ID проекта.
        sheet_name (str): Имя листа Excel.
        max_row (Optional[int]): Максимальный номер строки.
        max_column (Optional[int]): Максимальный номер столбца.

    Returns:
        Optional[int]: sheet_id листа или None в случае ошибки.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения листа.")
        return None

    try:
        cursor = connection.cursor()
        # Проверяем, существует ли лист
        cursor.execute(
            "SELECT sheet_id FROM sheets WHERE project_id = ? AND name = ?",
            (project_id, sheet_name)
        )
        result = cursor.fetchone()
        if result:
            logger.debug(f"Лист '{sheet_name}' уже существует с ID {result[0]}.")
            return result[0]
        else:
            # Создаем новый лист
            cursor.execute(
                "INSERT INTO sheets (project_id, name, max_row, max_column) VALUES (?, ?, ?, ?)",
                (project_id, sheet_name, max_row, max_column)
            )
            connection.commit()
            new_sheet_id = cursor.lastrowid
            logger.info(f"Создан новый лист '{sheet_name}' с ID {new_sheet_id}.")
            return new_sheet_id
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении листа '{sheet_name}': {e}")
        return None
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении листа '{sheet_name}': {e}", exc_info=True)
        return None


def load_all_sheets_metadata(connection: sqlite3.Connection, project_id: int = 1) -> List[Dict[str, Any]]:
    """
    Загружает метаданные (ID и имя) для всех листов в проекте.
    Используется для экспорта, чтобы знать, какие листы обрабатывать.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        project_id (int): ID проекта (по умолчанию 1 для MVP).

    Returns:
        List[Dict[str, Any]]: Список словарей с ключами 'sheet_id' и 'name'.
        Возвращает пустой список в случае ошибки или отсутствия листов.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки списка листов.")
        return [] # Возвращаем пустой список

    try:
        cursor = connection.cursor()
        # Выбираем sheet_id и name из таблицы sheets для заданного project_id
        # ORDER BY sheet_id для определённого порядка (можно изменить на name или добавить sheet_index в будущем)
        cursor.execute(
            "SELECT sheet_id, name, max_row, max_column FROM sheets WHERE project_id = ? ORDER BY sheet_id",
            (project_id,)
        )
        rows = cursor.fetchall()
        
        sheets_list = []
        for row in rows:
            sheet_info = {
                'sheet_id': row[0],
                'name': row[1],
                'max_row': row[2],
                'max_column': row[3],
            }
            sheets_list.append(sheet_info)
        
        logger.info(f"Загружено {len(sheets_list)} листов для проекта ID {project_id}.")
        return sheets_list

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке списка листов для проекта ID {project_id}: {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке списка листов для проекта ID {project_id}: {e}", exc_info=True)
        return []


def load_sheet_by_name(connection: sqlite3.Connection, project_id: int, sheet_name: str) -> Optional[Dict[str, Any]]:
    """
    Загружает метаданные (ID, имя, max_row, max_column) для конкретного листа по имени.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        project_id (int): ID проекта.
        sheet_name (str): Имя листа Excel.

    Returns:
        Optional[Dict[str, Any]]: Словарь с 'sheet_id', 'name', 'max_row', 'max_column' или None.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки листа по имени.")
        return None

    try:
        cursor = connection.cursor()
        cursor.execute(
            "SELECT sheet_id, name, max_row, max_column FROM sheets WHERE project_id = ? AND name = ?",
            (project_id, sheet_name)
        )
        row = cursor.fetchone()
        
        if row:
            return {
                'sheet_id': row[0],
                'name': row[1],
                'max_row': row[2],
                'max_column': row[3],
            }
        else:
            logger.info(f"Лист '{sheet_name}' не найден в проекте ID {project_id}.")
            return None

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке листа '{sheet_name}' из проекта ID {project_id}: {e}")
        return None
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке листа '{sheet_name}' из проекта ID {project_id}: {e}", exc_info=True)
        return None

# --- НОВОЕ: Функция для переименования листа ---
def rename_sheet(connection: sqlite3.Connection, project_id: int, old_name: str, new_name: str) -> bool:
    """
    Переименовывает лист в таблице 'sheets'.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        project_id (int): ID проекта.
        old_name (str): Текущее имя листа.
        new_name (str): Новое имя листа.

    Returns:
        bool: True, если переименование успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для переименования листа.")
        return False

    # Проверим, что новый лист с таким именем не существует
    try:
        cursor = connection.cursor()
        cursor.execute(
            "SELECT sheet_id FROM sheets WHERE project_id = ? AND name = ?",
            (project_id, new_name)
        )
        result = cursor.fetchone()
        if result:
            logger.warning(f"Лист с именем '{new_name}' уже существует в проекте ID {project_id}. Переименование невозможно.")
            return False
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при проверке существования листа '{new_name}' для переименования: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при проверке существования листа '{new_name}' для переименования: {e}", exc_info=True)
        return False

    try:
        cursor = connection.cursor()
        cursor.execute(
            "UPDATE sheets SET name = ? WHERE project_id = ? AND name = ?",
            (new_name, project_id, old_name)
        )
        connection.commit()
        if cursor.rowcount > 0:
            logger.info(f"Лист '{old_name}' успешно переименован в '{new_name}' в проекте ID {project_id}.")
            return True
        else:
            logger.warning(f"Лист '{old_name}' не найден в проекте ID {project_id} для переименования.")
            return False
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при переименовании листа '{old_name}' в '{new_name}' для проекта ID {project_id}: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при переименовании листа '{old_name}' в '{new_name}' для проекта ID {project_id}: {e}", exc_info=True)
        return False
# --- КОНЕЦ НОВОГО ---