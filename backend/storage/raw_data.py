# src/storage/raw_data.py

import sqlite3
import logging
from typing import List, Dict, Any
import re
from datetime import datetime

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

# --- ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ ФОРМАТИРОВАНИЯ ДАТЫ ДЛЯ GUI ---

def _format_datetime_for_gui(dt_str: str) -> str:
    """
    Преобразует строку даты из формата 'YYYY-MM-DD HH:MM:SS' в 'DD.MM.YYYY'.
    Если строка не соответствует формату, возвращается исходная строка.
    """
    try:
        # Пытаемся распарсить строку как datetime
        dt_obj = datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
        # Форматируем в нужный формат для GUI
        return dt_obj.strftime("%d.%m.%Y")
    except ValueError:
        # Если не удалось распарсить, возвращаем исходную строку
        logger.debug(f"Не удалось преобразовать строку '{dt_str}' в дату для GUI. Возвращается исходная строка.")
        return dt_str

# --- КОНЕЦ ВСПОМОГАТЕЛЬНОЙ ФУНКЦИИ ---

def _get_raw_data_table_name(sheet_name: str) -> str:
    """
    Генерирует имя таблицы для "сырых" данных листа.
    Args:
        sheet_name (str): Имя листа Excel.
    Returns:
        str: Имя таблицы в БД.
    """
    # Санитизация имени таблицы для безопасности и корректности SQL
    # Разрешаем только буквы, цифры и подчеркивания, заменяем пробелы подчеркиваниями
    sanitized_sheet_name = re.sub(r'[^\w]', '_', sheet_name)
    # Убедимся, что имя не начинается с цифры
    if sanitized_sheet_name and sanitized_sheet_name[0].isdigit():
        sanitized_sheet_name = f"_{sanitized_sheet_name}"
    # Ограничиваем длину имени таблицы (SQLite ограничение, обычно 64 символа)
    # Оставляем место для префикса 'raw_data_'
    max_len = 50 # Примерное ограничение
    if len(sanitized_sheet_name) > max_len:
        sanitized_sheet_name = sanitized_sheet_name[:max_len]
        
    return f"raw_data_{sanitized_sheet_name}"


def save_sheet_raw_data(connection: sqlite3.Connection, sheet_name: str, raw_data_list: List[Dict[str, Any]]) -> bool:
    """
    Сохраняет "сырые" данные листа в БД проекта.
    Создает отдельную таблицу для данных листа, если она не существует.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_name (str): Имя листа Excel.
        raw_data_list (List[Dict[str, Any]]): Список словарей с 'cell_address', 'value', 'value_type'.

    Returns:
        bool: True, если сохранение успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения сырых данных.")
        return False

    if not isinstance(raw_data_list, list):
        logger.error(f"Неверный тип данных для raw_data_list. Ожидался list, получен {type(raw_data_list)}.")
        return False

    try:
        cursor = connection.cursor()
        table_name = _get_raw_data_table_name(sheet_name)

        # Создаем таблицу для сырых данных листа, если она не существует
        logger.debug(f"Создание/проверка таблицы '{table_name}' для сырых данных листа '{sheet_name}'...")
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {table_name} (
                cell_address TEXT PRIMARY KEY,
                value TEXT,
                value_type TEXT
            )
        """)
        logger.debug(f"Таблица '{table_name}' готова.")

        # Подготавливаем данные для вставки/обновления
        # Используем INSERT OR REPLACE для простоты и атомарности
        data_to_insert = [
            (item.get('cell_address'), item.get('value'), type(item.get('value')).__name__)
            for item in raw_data_list
            if item.get('cell_address') # Пропускаем записи без адреса
        ]

        if data_to_insert:
            logger.debug(f"Подготовлено {len(data_to_insert)} записей сырых данных для листа '{sheet_name}'.")
            cursor.executemany(
                f"INSERT OR REPLACE INTO {table_name} (cell_address, value, value_type) VALUES (?, ?, ?)",
                data_to_insert
            )
            connection.commit()
            logger.info(f"Сохранено {len(data_to_insert)} записей сырых данных для листа '{sheet_name}' в таблицу '{table_name}'.")
        else:
            logger.info(f"Нет сырых данных для сохранения для листа '{sheet_name}'.")
            
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении сырых данных для листа '{sheet_name}': {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении сырых данных для листа '{sheet_name}': {e}", exc_info=True)
        return False


def load_sheet_raw_data(connection: sqlite3.Connection, sheet_name: str) -> List[Dict[str, Any]]:
    """
    Загружает "сырые" данные листа из БД проекта.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_name (str): Имя листа Excel.

    Returns:
        List[Dict[str, Any]]: Список словарей с 'cell_address', 'value', 'value_type'.
                             Возвращает пустой список в случае ошибки или отсутствия данных/таблицы.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки сырых данных.")
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
            logger.info(f"Таблица сырых данных '{table_name}' для листа '{sheet_name}' не найдена. Возвращается пустой список.")
            return []

        # Загружаем данные
        cursor.execute(f"SELECT cell_address, value, value_type FROM {table_name}")
        rows = cursor.fetchall()
        
        raw_data = []
        for row in rows:
            cell_address, value, value_type = row
            # --- ИЗМЕНЕНИЕ: Форматируем дату для GUI ---
            if value_type == 'datetime' and isinstance(value, str):
                formatted_value = _format_datetime_for_gui(value)
                raw_data.append({"cell_address": cell_address, "value": formatted_value, "value_type": value_type})
            else:
                raw_data.append({"cell_address": cell_address, "value": value, "value_type": value_type})
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---
        
        logger.debug(f"Загружено {len(raw_data)} записей сырых данных для листа '{sheet_name}' из таблицы '{table_name}'.")
        return raw_data

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке сырых данных для листа '{sheet_name}': {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке сырых данных для листа '{sheet_name}': {e}", exc_info=True)
        return []

# Дополнительные функции для работы с сырыми данными (если потребуются) могут быть добавлены здесь
