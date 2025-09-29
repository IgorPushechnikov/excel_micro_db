# src/storage/styles.py

import sqlite3
import logging
from typing import List, Dict, Any, Optional
import json

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

def save_sheet_styles(connection: sqlite3.Connection, sheet_id: int, styles_list: List[Dict[str, Any]]) -> bool:
    """
    Сохраняет стили листа в БД проекта.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.
        styles_list (List[Dict[str, Any]]): Список словарей с 'style_attributes' и 'range_address'.
                                           'style_attributes' - это словарь или JSON-строка атрибутов стиля.
                                           'range_address' - это строка адреса диапазона (например, 'A1:B10').

    Returns:
        bool: True, если сохранение успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения стилей.")
        return False

    if not isinstance(styles_list, list):
        logger.error(f"Неверный тип данных для styles_list. Ожидался list, получен {type(styles_list)}.")
        return False

    try:
        cursor = connection.cursor()
        
        # Удаляем существующие стили для этого листа, чтобы избежать дубликатов
        # Предполагается, что стили хранятся в таблице 'sheet_styles'
        cursor.execute("DELETE FROM sheet_styles WHERE sheet_id = ?", (sheet_id,))
        logger.debug(f"Удалены существующие стили для sheet_id {sheet_id}.")

        # Подготавливаем данные для вставки
        styles_to_insert = []
        for style_data in styles_list:
            range_address = style_data.get('range_address')
            style_attributes = style_data.get('style_attributes')
            
            if not range_address:
                logger.warning("Найдена запись стиля без 'range_address'. Пропущена.")
                continue

            # Сериализуем атрибуты стиля в JSON, если это словарь
            if isinstance(style_attributes, dict):
                try:
                    style_attributes_json = json.dumps(style_attributes, ensure_ascii=False)
                except (TypeError, ValueError) as e:
                    logger.error(f"Ошибка сериализации атрибутов стиля для диапазона {range_address}: {e}")
                    # Можно либо пропустить, либо сохранить как пустую строку/NULL
                    style_attributes_json = "{}" # Или None, в зависимости от схемы
            elif isinstance(style_attributes, str):
                 # Предполагаем, что это уже корректный JSON
                 style_attributes_json = style_attributes
            else:
                 logger.warning(f"Неподдерживаемый тип для 'style_attributes' в диапазоне {range_address}. Тип: {type(style_attributes)}. Пропущен.")
                 style_attributes_json = "{}" # Или None

            styles_to_insert.append((sheet_id, range_address, style_attributes_json))

        if styles_to_insert:
            # Вставляем новые стили
            # Предполагается, что таблица 'sheet_styles' имеет столбцы: sheet_id, range_address, style_attributes
            cursor.executemany(
                "INSERT INTO sheet_styles (sheet_id, range_address, style_attributes) VALUES (?, ?, ?)",
                styles_to_insert
            )
            connection.commit()
            logger.info(f"Сохранено {len(styles_to_insert)} стилей для листа ID {sheet_id}.")
        else:
             logger.info(f"Нет стилей для сохранения для листа ID {sheet_id}.")
             
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении стилей для листа ID {sheet_id}: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении стилей для листа ID {sheet_id}: {e}", exc_info=True)
        return False


def load_sheet_styles(connection: sqlite3.Connection, sheet_id: int) -> List[Dict[str, Any]]:
    """
    Загружает стили и диапазоны для указанного листа.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.

    Returns:
        List[Dict[str, Any]]: Список словарей с 'style_attributes' (строка JSON) и 'range_address'.
                             Возвращает пустой список в случае ошибки или отсутствия данных.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки стилей.")
        return []

    try:
        cursor = connection.cursor()
        
        # Загружаем стили для этого листа
        # Предполагается, что таблица 'sheet_styles' имеет столбцы: sheet_id, range_address, style_attributes
        cursor.execute(
            "SELECT range_address, style_attributes FROM sheet_styles WHERE sheet_id = ?",
            (sheet_id,)
        )
        rows = cursor.fetchall()
        
        styles_data = []
        for row in rows:
            range_address, style_attributes_json = row
            # style_attributes_json остается строкой JSON, как она хранится в БД
            styles_data.append({
                "range_address": range_address,
                "style_attributes": style_attributes_json # Экспортёр будет парсить это при необходимости
            })
            
        logger.debug(f"Загружено {len(styles_data)} записей стилей для листа ID {sheet_id}.")
        return styles_data

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке стилей для листа ID {sheet_id}: {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке стилей для листа ID {sheet_id}: {e}", exc_info=True)
        return []

# Дополнительные функции для работы со стилями (если потребуются) могут быть добавлены здесь
