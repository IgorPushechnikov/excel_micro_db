# src/storage/charts.py

import sqlite3
import logging
# from typing import List, Dict, Any, Optional
# import json # Понадобится, если мы будем сериализовать в JSON

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

def save_sheet_charts(connection: sqlite3.Connection, sheet_id: int, charts_list: list[dict]) -> bool:
    """
    Сохраняет диаграммы листа в БД проекта.
    В текущей реализации предполагается, что данные диаграммы уже подготовлены
    анализатором в сериализованном виде (например, JSON или XML).

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.
        charts_list (list[dict]): Список словарей с данными диаграмм.
                                  Каждый словарь должен содержать как минимум ключ 'chart_data',
                                  который представляет собой сериализованные данные диаграммы
                                  (например, строка JSON или XML).

    Returns:
        bool: True, если сохранение успешно, иначе False.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для сохранения диаграмм.")
        return False

    if not isinstance(charts_list, list):
        logger.error(f"Неверный тип данных для charts_list. Ожидался list, получен {type(charts_list)}.")
        return False

    try:
        cursor = connection.cursor()
        
        # Удаляем существующие диаграммы для этого листа, чтобы избежать дубликатов
        # Предполагается, что диаграммы хранятся в таблице 'sheet_charts'
        cursor.execute("DELETE FROM sheet_charts WHERE sheet_id = ?", (sheet_id,))
        logger.debug(f"Удалены существующие диаграммы для sheet_id {sheet_id}.")

        # Подготавливаем данные для вставки
        # Предполагаем, что каждый элемент charts_list имеет ключ 'chart_data'
        # и, возможно, другие метаданные, например, 'chart_id' или 'position'.
        charts_to_insert = []
        for chart_data in charts_list:
            # chart_data_str = chart_data.get('chart_data')
            chart_data_str = chart_data # Предполагаем, что сам словарь/объект уже сериализован или готов к хранению
            # В реальной реализации нужно будет решить, как именно хранить chart_data.
            # Если это словарь, его нужно сериализовать, например, в JSON.
            # chart_data_str = json.dumps(chart_data, ensure_ascii=False) 
            
            if chart_data_str is None:
                 logger.warning("Найдена запись диаграммы без 'chart_data'. Пропущена.")
                 continue

            # Можно добавить другие поля, если они есть и нужны
            # chart_id = chart_data.get('chart_id', None) # Если ID генерируется БД, можно не передавать
            # position = chart_data.get('position', None)
            
            # charts_to_insert.append((sheet_id, chart_id, chart_data_str, position))
            charts_to_insert.append((sheet_id, str(chart_data_str))) # Простая вставка, адаптировать под схему

        if charts_to_insert:
            # Вставляем новые диаграммы
            # Предполагается, что таблица 'sheet_charts' имеет столбцы: sheet_id, chart_data
            # Если есть другие поля, их нужно добавить в запрос
            cursor.executemany(
                # "INSERT INTO sheet_charts (sheet_id, chart_id, chart_data, position) VALUES (?, ?, ?, ?)",
                "INSERT INTO sheet_charts (sheet_id, chart_data) VALUES (?, ?)",
                charts_to_insert
            )
            connection.commit()
            logger.info(f"Сохранено {len(charts_to_insert)} диаграмм для листа ID {sheet_id}.")
        else:
             logger.info(f"Нет диаграмм для сохранения для листа ID {sheet_id}.")
             
        return True

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при сохранении диаграмм для листа ID {sheet_id}: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при сохранении диаграмм для листа ID {sheet_id}: {e}", exc_info=True)
        return False


def load_sheet_charts(connection: sqlite3.Connection, sheet_id: int) -> list[dict]:
    """
    Загружает диаграммы для указанного листа.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
        sheet_id (int): ID листа в БД.

    Returns:
        list[dict]: Список словарей с данными диаграмм.
                    Каждый словарь содержит ключ 'chart_data' с сериализованными данными.
                    Возвращает пустой список в случае ошибки или отсутствия данных.
    """
    if not connection:
        logger.error("Нет активного соединения с БД для загрузки диаграмм.")
        return []

    try:
        cursor = connection.cursor()
        
        # Загружаем диаграммы для этого листа
        # Предполагается, что таблица 'sheet_charts' имеет столбцы: sheet_id, chart_data
        cursor.execute(
            # "SELECT chart_id, chart_data, position FROM sheet_charts WHERE sheet_id = ?",
            "SELECT chart_data FROM sheet_charts WHERE sheet_id = ?",
            (sheet_id,)
        )
        rows = cursor.fetchall()
        
        charts_data = []
        for row in rows:
            # chart_id, chart_data_str, position = row
            chart_data_str = row[0]
            # chart_data_str остается сериализованным, как оно хранится в БД
            # Экспортёр будет отвечать за десериализацию/использование
            charts_data.append({
                # "chart_id": chart_id,
                "chart_data": chart_data_str, # Это может быть строка JSON/XML или другой сериализованный формат
                # "position": position
            })
            
        logger.debug(f"Загружено {len(charts_data)} записей диаграмм для листа ID {sheet_id}.")
        return charts_data

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при загрузке диаграмм для листа ID {sheet_id}: {e}")
        return []
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке диаграмм для листа ID {sheet_id}: {e}", exc_info=True)
        return []

# Дополнительные функции для работы с диаграммами (если потребуются) могут быть добавлены здесь
