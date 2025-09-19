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

CREATE_CHART_AXES_TABLE = '''
    CREATE TABLE IF NOT EXISTS chart_axes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chart_id INTEGER NOT NULL,
        axis_type TEXT NOT NULL, -- 'x_axis', 'y_axis', 'z_axis'
        ax_id INTEGER, -- Идентификатор оси
        ax_pos TEXT, -- Положение оси
        delete_axis INTEGER, -- Удаление оси (BOOLEAN)
        title TEXT, -- Заголовок оси
        num_fmt TEXT, -- Формат чисел
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

CREATE_CHART_SERIES_TABLE = '''
    CREATE TABLE IF NOT EXISTS chart_series (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chart_id INTEGER NOT NULL,
        idx INTEGER NOT NULL, -- Индекс серии
        "order" INTEGER NOT NULL, -- Порядок серии
        tx TEXT, -- Заголовок серии
        shape TEXT, -- Форма маркеров
        smooth INTEGER, -- Сглаживание (BOOLEAN)
        invert_if_negative INTEGER, -- Инвертировать цвет (BOOLEAN)
        FOREIGN KEY (chart_id) REFERENCES charts (id)
    )
'''

CREATE_CHART_DATA_SOURCES_TABLE = '''
    CREATE TABLE IF NOT EXISTS chart_data_sources (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        series_id INTEGER NOT NULL, -- Ссылка на chart_series.id
        data_type TEXT NOT NULL, -- 'values', 'categories'
        formula TEXT NOT NULL, -- Строка формулы диапазона
        FOREIGN KEY (series_id) REFERENCES chart_series (id)
    )
'''

# - НОВЫЕ ТАБЛИЦЫ ДЛЯ СЫРЫХ ДАННЫХ -
CREATE_RAW_DATA_TABLES_REGISTRY = '''
    CREATE TABLE IF NOT EXISTS raw_data_tables_registry (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sheet_id INTEGER NOT NULL UNIQUE, -- Связь с таблицей sheets
        table_name TEXT NOT NULL UNIQUE,  -- Имя таблицы с сырыми данными
        FOREIGN KEY (sheet_id) REFERENCES sheets (id)
    )
'''
# - КОНЕЦ НОВЫХ ТАБЛИЦ -

# === НОВЫЕ ТАБЛИЦЫ ДЛЯ СТИЛЕЙ ===
CREATE_FONTS_TABLE = '''
    CREATE TABLE IF NOT EXISTS fonts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        sz REAL,
        b INTEGER, -- BOOLEAN
        i INTEGER, -- BOOLEAN
        u TEXT,
        strike INTEGER, -- BOOLEAN
        color TEXT,
        color_theme INTEGER,
        color_tint REAL,
        vert_align TEXT,
        scheme TEXT,
        UNIQUE(name, sz, b, i, u, strike, color, color_theme, color_tint, vert_align, scheme)
    )
'''

CREATE_PATTERN_FILLS_TABLE = '''
    CREATE TABLE IF NOT EXISTS pattern_fills (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        pattern_type TEXT,
        fg_color TEXT,
        fg_color_theme INTEGER,
        fg_color_tint REAL,
        bg_color TEXT,
        bg_color_theme INTEGER,
        bg_color_tint REAL,
        UNIQUE(pattern_type, fg_color, fg_color_theme, fg_color_tint, bg_color, bg_color_theme, bg_color_tint)
    )
'''

CREATE_BORDERS_TABLE = '''
    CREATE TABLE IF NOT EXISTS borders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        left_style TEXT,
        left_color TEXT,
        right_style TEXT,
        right_color TEXT,
        top_style TEXT,
        top_color TEXT,
        bottom_style TEXT,
        bottom_color TEXT,
        diagonal_style TEXT,
        diagonal_color TEXT,
        diagonal_up INTEGER, -- BOOLEAN
        diagonal_down INTEGER, -- BOOLEAN
        outline INTEGER, -- BOOLEAN
        UNIQUE(left_style, left_color, right_style, right_color, top_style, top_color,
               bottom_style, bottom_color, diagonal_style, diagonal_color,
               diagonal_up, diagonal_down, outline)
    )
'''

CREATE_ALIGNMENTS_TABLE = '''
    CREATE TABLE IF NOT EXISTS alignments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        horizontal TEXT,
        vertical TEXT,
        text_rotation INTEGER,
        wrap_text INTEGER, -- BOOLEAN
        shrink_to_fit INTEGER, -- BOOLEAN
        indent INTEGER,
        relative_indent INTEGER,
        justify_last_line INTEGER, -- BOOLEAN
        reading_order INTEGER,
        text_direction TEXT,
        UNIQUE(horizontal, vertical, text_rotation, wrap_text, shrink_to_fit,
               indent, relative_indent, justify_last_line, reading_order, text_direction)
    )
'''

CREATE_PROTECTIONS_TABLE = '''
    CREATE TABLE IF NOT EXISTS protections (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        locked INTEGER, -- BOOLEAN
        hidden INTEGER, -- BOOLEAN
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
# === КОНЕЦ НОВОЙ ТАБЛИЦЫ ===

# Список всех SQL-запросов
TABLE_CREATION_QUERIES = [
    CREATE_PROJECTS_TABLE,
    CREATE_SHEETS_TABLE,
    CREATE_FORMULAS_TABLE,
    CREATE_CROSS_SHEET_REFS_TABLE,
    CREATE_CHARTS_TABLE,
    CREATE_CHART_AXES_TABLE,
    CREATE_CHART_SERIES_TABLE,
    CREATE_CHART_DATA_SOURCES_TABLE,
    CREATE_RAW_DATA_TABLES_REGISTRY,
    CREATE_FONTS_TABLE,
    CREATE_PATTERN_FILLS_TABLE,
    CREATE_BORDERS_TABLE,
    CREATE_ALIGNMENTS_TABLE,
    CREATE_PROTECTIONS_TABLE,
    CREATE_CELL_STYLES_TABLE,
    CREATE_STYLED_RANGES_TABLE,
    CREATE_MERGED_CELLS_TABLE,
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
        raise # Повторно вызываем исключение, чтобы его мог обработать вызывающий код
