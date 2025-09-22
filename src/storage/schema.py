# src/storage/schema.py
"""
Модуль для инициализации схемы базы данных проекта Excel Micro DB.

Содержит SQL-запросы для создания таблиц и функцию для их выполнения.
"""

import sqlite3
import logging

logger = logging.getLogger(__name__)

# === ИЗМЕНЕНО: Экранирование имени столбца "references" ===
# SQL-запросы для создания таблиц

CREATE_PROJECTS_TABLE = '''
CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL, -- Храним как текст
    description TEXT
)
'''

CREATE_SHEETS_TABLE = '''
CREATE TABLE IF NOT EXISTS sheets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    sheet_index INTEGER NOT NULL, -- Порядковый номер листа
    structure TEXT, -- JSON-строка с описанием структуры
    raw_data_info TEXT, -- НОВОЕ: JSON-строка с информацией о сырых данных (имена столбцов)
    FOREIGN KEY (project_id) REFERENCES projects (id)
)
'''

CREATE_FORMULAS_TABLE = '''
CREATE TABLE IF NOT EXISTS formulas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL,
    cell TEXT NOT NULL,
    formula TEXT NOT NULL,
    "references" TEXT, -- Имя столбца в кавычках, так как 'references' является ключевым словом SQL
    FOREIGN KEY (sheet_id) REFERENCES sheets (id)
)
'''

CREATE_CROSS_SHEET_REFS_TABLE = '''
CREATE TABLE IF NOT EXISTS cross_sheet_references (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL, -- ID листа, где находится формула
    from_cell TEXT NOT NULL,
    from_formula TEXT NOT NULL,
    to_sheet TEXT NOT NULL,
    reference_type TEXT NOT NULL,
    reference_address TEXT NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets (id)
)
'''

# === ИЗМЕНЕНО: Структурированное хранение диаграмм ===

CREATE_CHARTS_TABLE = '''
CREATE TABLE IF NOT EXISTS charts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL,
    type TEXT NOT NULL, -- Тип диаграммы, например, 'BarChart'
    title TEXT, -- Заголовок диаграммы
    top_left_cell TEXT, -- Адрес верхней левой ячейки (например, 'A1')
    width REAL, -- Ширина диаграммы
    height REAL, -- Высота диаграммы
    style INTEGER, -- Стиль диаграммы openpyxl
    legend_position TEXT, -- Положение легенды
    auto_scaling INTEGER, -- Автоматическое масштабирование (BOOLEAN)
    plot_vis_only INTEGER, -- Отображать только видимые ячейки (BOOLEAN)
    FOREIGN KEY (sheet_id) REFERENCES sheets (id)
)
'''

# === ИСПРАВЛЕНО: Имена столбцов и добавлены недостающие ===
CREATE_CHART_AXES_TABLE = '''
CREATE TABLE IF NOT EXISTS chart_axes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chart_id INTEGER NOT NULL,
    axis_type TEXT NOT NULL, -- 'x_axis', 'y_axis', 'z_axis'
    ax_id INTEGER, -- Идентификатор оси
    ax_pos TEXT, -- Положение оси
    delete_axis INTEGER, -- Удаление оси (BOOLEAN) -- ИСПРАВЛЕНО: delete -> delete_axis
    title TEXT, -- Заголовок оси -- ИСПРАВЛЕНО: axis_title -> title
    num_fmt TEXT, -- Формат чисел -- ИСПРАВЛЕНО: number_format -> num_fmt
    major_tick_mark TEXT, -- Основная метка делений
    minor_tick_mark TEXT, -- Дополнительная метка делений
    tick_lbl_pos TEXT, -- Положение подписей делений
    crosses TEXT, -- Тип пересечения
    crosses_at REAL, -- Значение пересечения
    major_unit REAL, -- Основная единица
    minor_unit REAL, -- Дополнительная единица
    min REAL, -- Минимальное значение
    max REAL, -- Максимальное значение
    orientation TEXT, -- Ориентация
    log_base REAL, -- Основание логарифма
    major_gridlines INTEGER, -- Наличие основной сетки (BOOLEAN)
    minor_gridlines INTEGER, -- Наличие дополнительной сетки (BOOLEAN)
    FOREIGN KEY (chart_id) REFERENCES charts (id)
)
'''

# === ИСПРАВЛЕНО: Имена столбцов и добавлен idx ===
CREATE_CHART_SERIES_TABLE = '''
CREATE TABLE IF NOT EXISTS chart_series (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chart_id INTEGER NOT NULL,
    idx INTEGER NOT NULL, -- ИСПРАВЛЕНО: Добавлен недостающий столбец idx -- НОВОЕ
    "order" INTEGER NOT NULL, -- Порядок серии (зарезервированное слово SQL, используем кавычки)
    tx TEXT, -- Название серии
    shape TEXT, -- Форма маркера
    smooth INTEGER, -- Сглаживание линии (BOOLEAN)
    invert_if_negative INTEGER, -- Инвертировать цвет, если значение отрицательное (BOOLEAN)
    FOREIGN KEY (chart_id) REFERENCES charts (id)
)
'''

# === ИСПРАВЛЕНО: Имена столбцов ===
CREATE_CHART_DATA_SOURCES_TABLE = '''
CREATE TABLE IF NOT EXISTS chart_data_sources (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL, -- Лист, откуда берутся данные
    data_range TEXT NOT NULL, -- Диапазон ячеек, например, 'Sheet1!$A$1:$A$10'
    data_type TEXT NOT NULL, -- ИСПРАВЛЕНО: source_type -> data_type ('category', 'value', 'size' и т.д.)
    FOREIGN KEY (sheet_id) REFERENCES sheets (id)
)
'''

# === НОВАЯ ТАБЛИЦА ДЛЯ РЕГИСТРАЦИИ СЫРЫХ ДАННЫХ ===

CREATE_RAW_DATA_TABLES_REGISTRY = '''
CREATE TABLE IF NOT EXISTS raw_data_tables_registry (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL UNIQUE, -- Связь с листом
    table_name TEXT NOT NULL UNIQUE, -- Имя таблицы с сырыми данными
    FOREIGN KEY (sheet_id) REFERENCES sheets (id)
)
'''

# === НОВЫЕ ТАБЛИЦЫ ДЛЯ СТИЛЕЙ ===

# === ИСПРАВЛЕНО: Добавлены недостающие столбцы ===
CREATE_FONTS_TABLE = '''
CREATE TABLE IF NOT EXISTS fonts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    sz REAL, -- Размер
    b INTEGER, -- Жирный (BOOLEAN)
    i INTEGER, -- Курсив (BOOLEAN)
    u TEXT, -- Подчеркивание (например, 'single', 'double') -- НОВОЕ
    strike INTEGER, -- Зачеркнутый (BOOLEAN) -- НОВОЕ
    color_theme INTEGER,
    color_type TEXT,
    color_rgb TEXT,
    color_tint REAL, -- НОВОЕ: Добавлен недостающий столбец tint для цвета -- НОВОЕ
    vert_align TEXT, -- Вертикальное выравнивание текста в строке (например, 'superscript', 'subscript') -- НОВОЕ
    scheme TEXT, -- Схема шрифта -- НОВОЕ
    family INTEGER,
    charset INTEGER,
    -- ИЗМЕНЕНО: Добавлены новые столбцы в UNIQUE constraint
    UNIQUE(name, sz, b, i, u, strike, color_theme, color_type, color_rgb, color_tint, vert_align, scheme, family, charset)
)
'''

# === ИСПРАВЛЕНО: Добавлены недостающие столбцы tint ===
CREATE_PATTERN_FILLS_TABLE = '''
CREATE TABLE IF NOT EXISTS pattern_fills (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    patternType TEXT,
    fgColor_theme INTEGER,
    fgColor_type TEXT,
    fgColor_rgb TEXT,
    fgColor_tint REAL, -- НОВОЕ: Добавлен недостающий столбец tint для fg цвета -- НОВОЕ
    bgColor_theme INTEGER,
    bgColor_type TEXT,
    bgColor_rgb TEXT,
    bgColor_tint REAL, -- НОВОЕ: Добавлен недостающий столбец tint для bg цвета -- НОВОЕ
    -- ИЗМЕНЕНО: Добавлены новые столбцы в UNIQUE constraint
    UNIQUE(patternType, fgColor_theme, fgColor_type, fgColor_rgb, fgColor_tint, bgColor_theme, bgColor_type, bgColor_rgb, bgColor_tint)
)
'''

# === ИСПРАВЛЕНО: Добавлены недостающие столбцы tint ===
CREATE_BORDERS_TABLE = '''
CREATE TABLE IF NOT EXISTS borders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    left_style TEXT,
    left_color_theme INTEGER,
    left_color_type TEXT,
    left_color_rgb TEXT,
    left_color_tint REAL, -- НОВОЕ: Добавлен недостающий столбец tint для left цвета -- НОВОЕ
    right_style TEXT,
    right_color_theme INTEGER,
    right_color_type TEXT,
    right_color_rgb TEXT,
    right_color_tint REAL, -- НОВОЕ
    top_style TEXT,
    top_color_theme INTEGER,
    top_color_type TEXT,
    top_color_rgb TEXT,
    top_color_tint REAL, -- НОВОЕ
    bottom_style TEXT,
    bottom_color_theme INTEGER,
    bottom_color_type TEXT,
    bottom_color_rgb TEXT,
    bottom_color_tint REAL, -- НОВОЕ
    diagonal_style TEXT,
    diagonal_color_theme INTEGER,
    diagonal_color_type TEXT,
    diagonal_color_rgb TEXT,
    diagonal_color_tint REAL, -- НОВОЕ
    diagonalUp INTEGER, -- BOOLEAN
    diagonalDown INTEGER, -- BOOLEAN
    outline INTEGER, -- BOOLEAN
    -- ИЗМЕНЕНО: Добавлены новые столбцы в UNIQUE constraint
    UNIQUE(left_style, left_color_theme, left_color_type, left_color_rgb, left_color_tint,
           right_style, right_color_theme, right_color_type, right_color_rgb, right_color_tint,
           top_style, top_color_theme, top_color_type, top_color_rgb, top_color_tint,
           bottom_style, bottom_color_theme, bottom_color_type, bottom_color_rgb, bottom_color_tint,
           diagonal_style, diagonal_color_theme, diagonal_color_type, diagonal_color_rgb, diagonal_color_tint,
           diagonalUp, diagonalDown, outline)
)
'''

CREATE_ALIGNMENTS_TABLE = '''
CREATE TABLE IF NOT EXISTS alignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    horizontal TEXT,
    vertical TEXT,
    textRotation REAL,
    wrapText INTEGER, -- BOOLEAN
    shrinkToFit INTEGER, -- BOOLEAN
    indent REAL,
    relativeIndent REAL,
    justifyLastLine INTEGER, -- BOOLEAN
    readingOrder REAL,
    UNIQUE(horizontal, vertical, textRotation, wrapText, shrinkToFit, indent, relativeIndent, justifyLastLine, readingOrder)
)
'''

# === ИСПРАВЛЕНО: Добавлены значения по умолчанию ===
CREATE_PROTECTIONS_TABLE = '''
CREATE TABLE IF NOT EXISTS protections (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    locked INTEGER DEFAULT 1, -- BOOLEAN, по умолчанию заблокировано -- ИЗМЕНЕНО
    hidden INTEGER DEFAULT 0, -- BOOLEAN, по умолчанию не скрыто -- ИЗМЕНЕНО
    UNIQUE(locked, hidden)
)
'''

CREATE_CELL_STYLES_TABLE = '''
CREATE TABLE IF NOT EXISTS cell_styles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    font_id INTEGER REFERENCES fonts(id),
    fill_id INTEGER REFERENCES pattern_fills(id),
    border_id INTEGER REFERENCES borders(id),
    alignment_id INTEGER REFERENCES alignments(id),
    protection_id INTEGER REFERENCES protections(id),
    num_fmt_id INTEGER, -- ID формата чисел
    xf_id INTEGER, -- ID XF
    quote_prefix INTEGER, -- BOOLEAN
    UNIQUE(font_id, fill_id, border_id, alignment_id, protection_id, num_fmt_id, xf_id, quote_prefix)
)
'''

CREATE_STYLED_RANGES_TABLE = '''
CREATE TABLE IF NOT EXISTS styled_ranges (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL,
    style_id INTEGER NOT NULL, -- Ссылка на cell_styles.id
    range_address TEXT NOT NULL, -- Адрес диапазона (например, 'A1:B10')
    FOREIGN KEY (sheet_id) REFERENCES sheets (id),
    FOREIGN KEY (style_id) REFERENCES cell_styles (id)
)
'''

# === КОНЕЦ НОВЫХ ТАБЛИЦ ДЛЯ СТИЛЕЙ ===

# === НОВАЯ ТАБЛИЦА ДЛЯ ОБЪЕДИНЕННЫХ ЯЧЕЕК ===

CREATE_MERGED_CELLS_TABLE = '''
CREATE TABLE IF NOT EXISTS merged_cells_ranges (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL,
    range_address TEXT NOT NULL, -- Адрес объединенного диапазона (например, 'A1:C3')
    FOREIGN KEY (sheet_id) REFERENCES sheets (id)
)
'''

# === НОВАЯ ТАБЛИЦА ДЛЯ ИСТОРИИ РЕДАКТИРОВАНИЯ ===

CREATE_EDIT_HISTORY_TABLE = '''
CREATE TABLE IF NOT EXISTS edit_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    sheet_id INTEGER,
    cell_address TEXT, -- Можно также хранить row_index и column_name отдельно
    action_type TEXT NOT NULL, -- 'edit_cell', 'add_row', 'delete_row', 'apply_style' и т.д.
    old_value TEXT, -- Старое значение ячейки
    new_value TEXT, -- Новое значение ячейки
    timestamp TEXT NOT NULL, -- Время изменения в ISO формате
    user TEXT, -- Имя пользователя (если есть система пользователей)
    details TEXT, -- Дополнительная информация в формате JSON
    FOREIGN KEY (project_id) REFERENCES projects (id),
    FOREIGN KEY (sheet_id) REFERENCES sheets (id)
)
'''

# ================================================

# Список всех SQL-запросов - ДОБАВЛЕН НОВЫЙ ЗАПРОС В КОНЕЦ СПИСКА
TABLE_CREATION_QUERIES = [\
    CREATE_PROJECTS_TABLE,\
    CREATE_SHEETS_TABLE,\
    CREATE_FORMULAS_TABLE,\
    CREATE_CROSS_SHEET_REFS_TABLE,\
    CREATE_CHARTS_TABLE,\
    CREATE_CHART_AXES_TABLE,\
    CREATE_CHART_SERIES_TABLE,\
    CREATE_CHART_DATA_SOURCES_TABLE,\
    CREATE_RAW_DATA_TABLES_REGISTRY,\
    CREATE_FONTS_TABLE,\
    CREATE_PATTERN_FILLS_TABLE,\
    CREATE_BORDERS_TABLE,\
    CREATE_ALIGNMENTS_TABLE,\
    CREATE_PROTECTIONS_TABLE,\
    CREATE_CELL_STYLES_TABLE,\
    CREATE_STYLED_RANGES_TABLE,\
    CREATE_MERGED_CELLS_TABLE,\
    CREATE_EDIT_HISTORY_TABLE, # <-- ДОБАВИЛИ СЮДА\
]

def initialize_schema(cursor: sqlite3.Cursor):
    """
    Инициализирует схему базы данных, создавая необходимые таблицы,
    если они еще не существуют.

    Args:
        cursor (sqlite3.Cursor): Курсор для выполнения SQL-запросов.
    """
    if not cursor:
        logger.error("Получен пустой курсор для инициализации схемы.")
        return

    try:
        for query in TABLE_CREATION_QUERIES:
            cursor.execute(query)
        logger.info("Схема базы данных инициализирована успешно.")
    except sqlite3.Error as e:
        logger.error(f"Ошибка SQLite при инициализации схемы: {e}")
        raise  # Повторно вызываем исключение, чтобы его мог обработать вызывающий код
