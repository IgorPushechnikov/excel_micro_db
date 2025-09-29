# src/storage/schema.py

import sqlite3
import logging

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

# SQL-запросы для создания таблиц проекта

# --- Таблицы для управления проектами ---

SQL_CREATE_PROJECTS_TABLE = """
CREATE TABLE IF NOT EXISTS projects (
    project_id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    description TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    last_opened_at TEXT
);
"""

# --- Таблицы для хранения информации о листах ---

SQL_CREATE_SHEETS_TABLE = """
CREATE TABLE IF NOT EXISTS sheets (
    sheet_id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    max_row INTEGER,
    max_column INTEGER,
    FOREIGN KEY (project_id) REFERENCES projects (project_id) ON DELETE CASCADE
);
"""

# --- Таблицы для хранения "сырых" данных ---
# Таблица для хранения "сырых" данных будет создаваться динамически для каждого листа
# как raw_data_<sanitized_sheet_name>

# --- Таблицы для хранения редактируемых данных ---
# Таблица для хранения редактируемых данных будет создаваться динамически для каждого листа
# как editable_data_<sanitized_sheet_name>

# --- Таблицы для хранения формул ---

SQL_CREATE_FORMULAS_TABLE = """
CREATE TABLE IF NOT EXISTS formulas (
    formula_id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL,
    cell_address TEXT NOT NULL,
    formula TEXT NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets (sheet_id) ON DELETE CASCADE,
    UNIQUE(sheet_id, cell_address)
);
"""

# --- Таблицы для хранения стилей ---

# Таблица для хранения определений уникальных стилей
SQL_CREATE_STYLES_TABLE = """
CREATE TABLE IF NOT EXISTS styles (
    style_id INTEGER PRIMARY KEY AUTOINCREMENT
    -- Можно добавить общие атрибуты стиля, если нужно для поиска/индексации
    -- Например: name TEXT UNIQUE
    -- Пока храним всё в sheet_styles
);
"""

# Таблица для связывания стилей с диапазонами на листах
# Хранит сериализованные атрибуты стиля (например, JSON)
SQL_CREATE_SHEET_STYLES_TABLE = """
CREATE TABLE IF NOT EXISTS sheet_styles (
    sheet_id INTEGER NOT NULL,
    range_address TEXT NOT NULL, -- Адрес диапазона, например, "A1:B10"
    style_attributes TEXT NOT NULL, -- Сериализованные атрибуты стиля (например, JSON)
    -- Если style_id используется, можно добавить FOREIGN KEY
    -- style_id INTEGER,
    -- FOREIGN KEY (sheet_id) REFERENCES sheets (sheet_id) ON DELETE CASCADE,
    -- FOREIGN KEY (style_id) REFERENCES styles (style_id) ON DELETE CASCADE,
    PRIMARY KEY (sheet_id, range_address)
    -- Если один диапазон может иметь несколько стилей, нужно убрать PRIMARY KEY
    -- и сделать отдельную таблицу связей.
);
"""

# --- Таблицы для хранения диаграмм ---

# Таблица для хранения данных диаграмм
# Хранит сериализованные данные диаграммы (например, JSON или XML)
SQL_CREATE_SHEET_CHARTS_TABLE = """
CREATE TABLE IF NOT EXISTS sheet_charts (
    -- chart_id INTEGER PRIMARY KEY AUTOINCREMENT, -- Можно генерировать ID в БД
    sheet_id INTEGER NOT NULL,
    -- Другие метаданные диаграммы
    -- chart_type TEXT,
    -- position TEXT, -- JSON с позицией и размером
    chart_data TEXT NOT NULL, -- Сериализованные данные диаграммы (JSON/XML/BLOB)
    FOREIGN KEY (sheet_id) REFERENCES sheets (sheet_id) ON DELETE CASCADE
    -- PRIMARY KEY (sheet_id, chart_data) -- Или другой уникальный ключ
);
"""

# --- Таблицы для хранения объединенных ячеек ---

# Таблица для хранения данных об объединенных ячейках на листе
# Хранит сериализованный JSON-массив строк адресов диапазонов
SQL_CREATE_SHEET_MERGED_CELLS_TABLE = """
CREATE TABLE IF NOT EXISTS sheet_merged_cells (
    sheet_id INTEGER NOT NULL,
    merged_cells_data TEXT NOT NULL, -- Сериализованный JSON-массив строк адресов диапазонов, например, '["A1:B2", "C3:D5"]'
    FOREIGN KEY (sheet_id) REFERENCES sheets (sheet_id) ON DELETE CASCADE,
    PRIMARY KEY (sheet_id)
);
"""

# --- Таблицы для хранения истории редактирования ---

# Исправленный SQL-запрос для создания таблицы истории редактирования
# Включает все необходимые поля
SQL_CREATE_EDIT_HISTORY_TABLE = """
CREATE TABLE IF NOT EXISTS edit_history (
    history_id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    sheet_id INTEGER NOT NULL,
    cell_address TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT,
    edited_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (project_id) REFERENCES projects (project_id) ON DELETE CASCADE,
    FOREIGN KEY (sheet_id) REFERENCES sheets (sheet_id) ON DELETE CASCADE
);
"""

# --- Таблицы для хранения метаданных проекта ---

# Для хранения произвольных пар ключ-значение для проекта
SQL_CREATE_PROJECT_METADATA_TABLE = """
CREATE TABLE IF NOT EXISTS project_metadata (
    project_id INTEGER NOT NULL,
    key TEXT NOT NULL,
    value TEXT,
    FOREIGN KEY (project_id) REFERENCES projects (project_id) ON DELETE CASCADE,
    PRIMARY KEY (project_id, key)
);
"""


def initialize_project_schema(connection: sqlite3.Connection):
    """
    Инициализирует схему таблиц проекта в БД.

    Args:
        connection (sqlite3.Connection): Активное соединение с БД проекта.
    """

    if not connection:
        logger.error("Нет активного соединения с БД для инициализации схемы.")
        return

    try:
        cursor = connection.cursor()

        # --- Создание таблиц ---

        logger.debug("Создание таблицы 'projects'...")
        cursor.execute(SQL_CREATE_PROJECTS_TABLE)

        logger.debug("Создание таблицы 'sheets'...")
        cursor.execute(SQL_CREATE_SHEETS_TABLE)

        logger.debug("Создание таблицы 'formulas'...")
        cursor.execute(SQL_CREATE_FORMULAS_TABLE)

        logger.debug("Создание таблицы 'styles'...")
        cursor.execute(SQL_CREATE_STYLES_TABLE)

        logger.debug("Создание таблицы 'sheet_styles'...")
        cursor.execute(SQL_CREATE_SHEET_STYLES_TABLE)

        logger.debug("Создание таблицы 'sheet_charts'...")
        cursor.execute(SQL_CREATE_SHEET_CHARTS_TABLE)

        logger.debug("Создание таблицы 'sheet_merged_cells'...")
        cursor.execute(SQL_CREATE_SHEET_MERGED_CELLS_TABLE)

        # --- Исправление: Создание таблицы 'edit_history' ---
        logger.debug("Создание таблицы 'edit_history'...")
        cursor.execute(SQL_CREATE_EDIT_HISTORY_TABLE)

        logger.debug("Создание таблицы 'project_metadata'...")
        cursor.execute(SQL_CREATE_PROJECT_METADATA_TABLE)

        # --- Создание индексов для оптимизации ---

        # Индекс для быстрого поиска листов по project_id
        logger.debug("Создание индекса для 'sheets.project_id'...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sheets_project_id ON sheets(project_id);")

        # Индекс для быстрого поиска формул по sheet_id
        logger.debug("Создание индекса для 'formulas.sheet_id'...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_formulas_sheet_id ON formulas(sheet_id);")

        # Индекс для быстрого поиска стилей по sheet_id
        logger.debug("Создание индекса для 'sheet_styles.sheet_id'...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sheet_styles_sheet_id ON sheet_styles(sheet_id);")

        # Индекс для быстрого поиска диаграмм по sheet_id
        logger.debug("Создание индекса для 'sheet_charts.sheet_id'...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sheet_charts_sheet_id ON sheet_charts(sheet_id);")

        # Индекс для быстрого поиска истории по project_id и sheet_id
        logger.debug("Создание индекса для 'edit_history.project_id'...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_edit_history_project_id ON edit_history(project_id);")

        logger.debug("Создание индекса для 'edit_history.sheet_id'...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_edit_history_sheet_id ON edit_history(sheet_id);")

        connection.commit()
        logger.debug("Commit выполнен. Проверка наличия таблиц...")
        # Принудительная проверка, что таблицы созданы
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables_after_commit = [row[0] for row in cursor.fetchall()]
        logger.debug(f"Таблицы после commit: {tables_after_commit}")
        if 'sheet_merged_cells' not in tables_after_commit:
            logger.error("КРИТИЧЕСКАЯ ОШИБКА: Таблица 'sheet_merged_cells' НЕ создана после commit!")
        logger.info("Схема таблиц проекта успешно инициализирована.")

    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при инициализации схемы: {e}")
        raise  # Повторно вызываем исключение, чтобы ошибка передалась выше
    except Exception as e:
        logger.error(f"Неожиданная ошибка при инициализации схемы: {e}", exc_info=True)
        raise  # Повторно вызываем исключение

# Дополнительные функции для работы со схемой (если потребуются) могут быть добавлены здесь