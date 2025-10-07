# src/storage/metadata.py

import sqlite3
import logging
from typing import Dict, Any, Optional
import json

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

# Имя таблицы для хранения метаданных проекта/листа
METADATA_TABLE_NAME = "project_metadata" # Используем общую таблицу, как определено в schema.py

def save_sheet_metadata(connection: sqlite3.Connection, sheet_name: str, metadata: Dict[str, Any]) -> bool:
    """
    Сохраняет метаданные листа в БД проекта.
    Метаданные хранятся в общей таблице 'project_metadata' с ключом 'sheet_<sheet_name>_<meta_key>'.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_name (str): Имя листа Excel.
        metadata (Dict[str, Any]): Словарь с метаданными листа (например, max_row, max_column, merged_cells).

    Returns:
        bool: True, если сохранение успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения метаданных листа.")
        return False

    if not isinstance(metadata, dict):
        logger.error(f"Неверный тип данных для metadata. Ожидался dict, получен {type(metadata)}.")
        return False

    try:
        cursor = connection.cursor()
        
        # Предполагаем, что project_id доступен или может быть получен.
        # В простейшем случае, если у нас один проект в БД, можно использовать project_id = 1.
        # Более правильно - получать project_id из контекста (например, через AppController).
        # Для MVP предположим, что project_id = 1.
        # TODO: Получать реальный project_id.
        project_id = 1 

        # Подготавливаем данные для вставки/обновления
        # Ключи метаданных листа будут иметь префикс 'sheet_<sheet_name>_'
        metadata_to_save = []
        for key, value in metadata.items():
            meta_key = f"sheet_{sheet_name}_{key}"
            # Сериализуем значение в JSON, если оно не является простым типом
            if isinstance(value, (dict, list)):
                try:
                    meta_value = json.dumps(value, ensure_ascii=False)
                except (TypeError, ValueError) as e:
                    logger.error(f"Ошибка сериализации метаданных '{meta_key}': {e}")
                    meta_value = str(value) # fallback
            else:
                meta_value = str(value) if value is not None else None
            
            metadata_to_save.append((project_id, meta_key, meta_value))

        if metadata_to_save:
            logger.debug(f"Подготовлено {len(metadata_to_save)} записей метаданных для листа '{sheet_name}'.")
            # Используем INSERT OR REPLACE для обновления существующих или вставки новых
            cursor.executemany(
                f"INSERT OR REPLACE INTO {METADATA_TABLE_NAME} (project_id, key, value) VALUES (?, ?, ?)",
                metadata_to_save
            )
            connection.commit()
            logger.info(f"Сохранено {len(metadata_to_save)} записей метаданных для листа '{sheet_name}' в таблицу '{METADATA_TABLE_NAME}'.")
        else:
            logger.info(f"Нет метаданных для сохранения для листа '{sheet_name}'.")
            
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении метаданных для листа '{sheet_name}': {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении метаданных для листа '{sheet_name}': {e}", exc_info=True)
        return False


def load_sheet_metadata(connection: sqlite3.Connection, sheet_name: str) -> Optional[Dict[str, Any]]:
    """
    Загружает метаданные для указанного листа из БД проекта.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_name (str): Имя листа Excel.

    Returns:
        Optional[Dict[str, Any]]: Словарь с метаданными листа или None в случае ошибки.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки метаданных листа.")
        return None

    try:
        cursor = connection.cursor()
        
        # Предполагаем project_id = 1 для MVP
        # TODO: Получать реальный project_id.
        project_id = 1
        
        # Загружаем метаданные, ключи которых начинаются с 'sheet_<sheet_name>_'
        prefix = f"sheet_{sheet_name}_"
        cursor.execute(
            f"SELECT key, value FROM {METADATA_TABLE_NAME} WHERE project_id = ? AND key LIKE ?",
            (project_id, f"{prefix}%")
        )
        rows = cursor.fetchall()
        
        sheet_metadata = {}
        for row in rows:
            key, value_str = row
            # Извлекаем оригинальный ключ метаданных (без префикса)
            original_key = key[len(prefix):]
            
            # Пытаемся десериализовать значение из JSON
            if value_str:
                try:
                    # Пробуем распарсить как JSON
                    sheet_metadata[original_key] = json.loads(value_str)
                except json.JSONDecodeError:
                    # Если не JSON, сохраняем как строку
                    sheet_metadata[original_key] = value_str
            else:
                sheet_metadata[original_key] = value_str # None или пустая строка
                
        logger.debug(f"Загружено {len(sheet_metadata)} записей метаданных для листа '{sheet_name}'.")
        return sheet_metadata if sheet_metadata else {}

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке метаданных для листа '{sheet_name}': {e}")
        return None
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке метаданных для листа '{sheet_name}': {e}", exc_info=True)
        return None

# --- НОВОЕ: Функция для сохранения метаданных проекта ---
def save_project_metadata(connection: sqlite3.Connection, project_id: int, metadata: Dict[str, Any]) -> bool:
    """
    Сохраняет метаданные проекта в БД.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД.
        project_id (int): ID проекта.
        metadata (Dict[str, Any]): Словарь с метаданными проекта.

    Returns:
        bool: True, если сохранение успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения метаданных проекта.")
        return False

    if not isinstance(metadata, dict):
        logger.error(f"Неверный тип данных для metadata. Ожидался dict, получен {type(metadata)}.")
        return False

    try:
        cursor = connection.cursor()

        # Подготавливаем данные для вставки/обновления
        metadata_to_save = []
        for key, value in metadata.items():
            # Сериализуем значение в JSON, если оно не является простым типом
            if isinstance(value, (dict, list)):
                try:
                    meta_value = json.dumps(value, ensure_ascii=False)
                except (TypeError, ValueError) as e:
                    logger.error(f"Ошибка сериализации метаданных проекта '{key}': {e}")
                    meta_value = str(value)  # fallback
            else:
                meta_value = str(value) if value is not None else None

            metadata_to_save.append((project_id, key, meta_value))

        if metadata_to_save:
            logger.debug(f"Подготовлено {len(metadata_to_save)} записей метаданных проекта.")
            # Используем INSERT OR REPLACE для обновления существующих или вставки новых
            cursor.executemany(
                f"INSERT OR REPLACE INTO {METADATA_TABLE_NAME} (project_id, key, value) VALUES (?, ?, ?)",
                metadata_to_save
            )
            connection.commit()
            logger.info(f"Сохранено {len(metadata_to_save)} записей метаданных проекта (ID: {project_id}) в таблицу '{METADATA_TABLE_NAME}'.")
        else:
            logger.info(f"Нет метаданных проекта для сохранения.")

        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении метаданных проекта (ID: {project_id}): {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении метаданных проекта (ID: {project_id}): {e}", exc_info=True)
        return False
# --- КОНЕЦ НОВОГО ---

# Дополнительные функции для работы с метаданными проекта (не только листа) могут быть добавлены здесь
