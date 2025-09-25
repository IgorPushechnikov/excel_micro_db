# src/utils/db_utils.py

import sqlite3
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

def dump_db_to_sql(db_path: str, output_sql_path: str):
    """
    Конвертирует базу данных SQLite в файл SQL-дампа.

    Args:
        db_path (str): Путь к файлу базы данных SQLite (.db).
        output_sql_path (str): Путь к выходному файлу SQL-дампа (.sql).
    """
    db_path = Path(db_path)
    output_sql_path = Path(output_sql_path)

    if not db_path.exists():
        logger.error(f"Файл БД не найден для дампа: {db_path}")
        return False

    try:
        # Открываем соединение с БД
        with sqlite3.connect(db_path) as conn:
            # Открываем файл для записи SQL
            with open(output_sql_path, 'w', encoding='utf-8') as f:
                # Используем iterdump для получения SQL-команд
                for line in conn.iterdump():
                    f.write(f"{line}\n")
        logger.info(f"Дамп БД успешно создан: {output_sql_path}")
        return True
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при создании дампа: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при создании дампа БД: {e}", exc_info=True)
        return False