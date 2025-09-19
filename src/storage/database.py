# src/storage/database.py
"""Модуль для работы с внутренним хранилищем данных проекта Excel Micro DB.
Использует SQLite для хранения данных и метаданных, извлеченных анализатором,
а также изменений, внесенных пользователем.
"""
import sys
import sqlite3
import json
import os
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import logging
# === ИСПРАВЛЕНО: Импорт класса datetime ===
from datetime import datetime # Импортируем класс datetime
# =========================================

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

# === ИСПРАВЛЕНО: Класс DateTimeEncoder ===
class DateTimeEncoder(json.JSONEncoder):
    """Пользовательский JSONEncoder для сериализации datetime объектов."""
    def default(self, obj):
        # Проверяем, является ли объект экземпляром класса datetime
        if isinstance(obj, datetime): # <-- ИСПРАВЛЕНО
            # Форматируем дату и время в строку в формате ISO 8601
            return obj.isoformat()
        # Для всех остальных типов вызываем метод родителя
        return super().default(obj)
# =========================================

def sanitize_table_name(name: str) -> str:
    """
    Санитизирует имя для использования в качестве имени таблицы SQLite.
    Заменяет недопустимые символы на подчеркивания и добавляет префикс, 
    если имя начинается с цифры.
    Args:
        name (str): Исходное имя.
    Returns:
        str: Санитизированное имя.
    """
    logger.debug(f"[DEBUG_STORAGE] Санитизация имени таблицы: '{name}'")
    if not name:
        logger.warning("[DEBUG_STORAGE] Получено пустое имя для санитизации. Возвращаю '_empty'.")
        return "_empty"
    
    # Заменяем все недопустимые символы (все, кроме букв, цифр и подчеркиваний) на '_'
    sanitized = "".join(ch if ch.isalnum() or ch == '_' else '_' for ch in name)
    
    # Если имя начинается с цифры, добавляем префикс
    if sanitized and sanitized[0].isdigit():
        sanitized = f"tbl_{sanitized}"
        logger.debug(f"[DEBUG_STORAGE] Имя начиналось с цифры. Добавлен префикс 'tbl_'. Новое имя: '{sanitized}'")
    
    # Убедимся, что имя не пустое и не состоит только из подчеркиваний
    if not sanitized or all(c == '_' for c in sanitized):
        sanitized = f"table_{abs(hash(name))}" # Создаем уникальное имя на основе хеша
        logger.debug(f"[DEBUG_STORAGE] Имя после санитации было пустым или некорректным. Создано новое имя: '{sanitized}'")
        
    logger.debug(f"[DEBUG_STORAGE] Санитизированное имя таблицы: '{sanitized}'")
    return sanitized

class ProjectDBStorage:
    """
    Класс для управления подключением к БД проекта и выполнения операций с данными.
    """
    def __init__(self, db_path: str):
        """
        Инициализирует объект хранилища.
        Args:
            db_path (str): Путь к файлу базы данных SQLite.
        """
        self.db_path = db_path
        self.connection: Optional[sqlite3.Connection] = None
        logger.debug(f"Инициализация хранилища проекта с БД: {self.db_path}")

    def connect(self):
        """Устанавливает соединение с базой данных."""
        try:
            # Убеждаемся, что директория для БД существует
            db_path_obj = Path(self.db_path)
            db_path_obj.parent.mkdir(parents=True, exist_ok=True)
            
            self.connection = sqlite3.connect(self.db_path)
            logger.debug(f"Соединение с БД {self.db_path} установлено.")
            # Инициализируем схему при подключении
            self._init_schema()
        except sqlite3.Error as e:
            logger.error(f"Ошибка подключения к БД {self.db_path}: {e}")
            raise

    def disconnect(self):
        """Закрывает соединение с базой данных."""
        if self.connection:
            self.connection.close()
            logger.debug("Соединение с БД закрыто.")
            self.connection = None

    def __enter__(self):
        """Поддержка контекстного менеджера 'with'."""
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Поддержка контекстного менеджера 'with'."""
        self.disconnect()

    def _init_schema(self):
        """
        Инициализирует схему базы данных, создавая необходимые таблицы,
        если они еще не существуют.
        """
        if not self.connection:
            logger.error("Нет активного соединения с БД для инициализации схемы.")
            return

        cursor = self.connection.cursor()

        # Создание таблицы проектов (если потребуется хранить метаинформацию о проекте)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL, -- Храним как текст
                description TEXT
            )
        ''')

        # Создание таблицы листов (ИЗМЕНЕНО: index -> sheet_index)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sheets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                sheet_index INTEGER NOT NULL, -- Порядковый номер листа
                structure TEXT, -- JSON-строка с описанием структуры
                raw_data_info TEXT, -- НОВОЕ: JSON-строка с информацией о сырых данных (имена столбцов)
                FOREIGN KEY (project_id) REFERENCES projects (id)
            )
        ''')

        # === ИЗМЕНЕНО: Экранирование имени столбца "references" ===
        # Создание таблицы формул
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS formulas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet_id INTEGER NOT NULL,
                cell TEXT NOT NULL,
                formula TEXT NOT NULL,
                "references" TEXT, -- Имя столбца в кавычках, так как 'references' является ключевым словом SQL
                FOREIGN KEY (sheet_id) REFERENCES sheets (id)
            )
        ''')
        # ===========================================================

        # Создание таблицы межлистовых ссылок
        cursor.execute('''
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
        ''')

        # === ИЗМЕНЕНО: Структурированное хранение диаграмм ===
        # Таблица для основной информации о диаграммах
        cursor.execute('''
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
        ''')

        # Таблица для осей диаграмм
        cursor.execute('''
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
        ''')

        # Таблица для серий данных диаграмм
        cursor.execute('''
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
        ''')

        # Таблица для источников данных серий
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS chart_data_sources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                series_id INTEGER NOT NULL, -- Ссылка на chart_series.id
                data_type TEXT NOT NULL, -- 'values', 'categories'
                formula TEXT NOT NULL, -- Строка формулы диапазона
                FOREIGN KEY (series_id) REFERENCES chart_series (id)
            )
        ''')
        # ===========================================================

        # - НОВЫЕ ТАБЛИЦЫ ДЛЯ СЫРЫХ ДАННЫХ -
        # Таблицы для сырых данных будут создаваться динамически для каждого листа
        # в методе create_raw_data_table. Здесь создаем таблицу для отслеживания этих таблиц.
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS raw_data_tables_registry (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet_id INTEGER NOT NULL UNIQUE, -- Связь с таблицей sheets
                table_name TEXT NOT NULL UNIQUE,  -- Имя таблицы с сырыми данными
                FOREIGN KEY (sheet_id) REFERENCES sheets (id)
            )
        ''')
        # - КОНЕЦ НОВЫХ ТАБЛИЦ -

        # === НОВЫЕ ТАБЛИЦЫ ДЛЯ СТИЛЕЙ ===
        # Таблица для шрифтов
        cursor.execute('''
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
        ''')
        
        # Таблица для заливок
        cursor.execute('''
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
        ''')

        # Таблица для границ
        cursor.execute('''
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
        ''')

        # Таблица для выравнивания
        cursor.execute('''
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
        ''')

        # Таблица для защиты
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS protections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                locked INTEGER, -- BOOLEAN
                hidden INTEGER, -- BOOLEAN
                UNIQUE(locked, hidden)
            )
        ''')

        # Таблица для уникальных определений стилей ячеек
        cursor.execute('''
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
        ''')

        # Таблица для хранения связи между стилями и диапазонами ячеек
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS styled_ranges (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet_id INTEGER NOT NULL,
                style_id INTEGER NOT NULL, -- Ссылка на cell_styles.id
                range_address TEXT NOT NULL, -- Адрес диапазона (например, 'A1:B10')
                FOREIGN KEY (sheet_id) REFERENCES sheets (id),
                FOREIGN KEY (style_id) REFERENCES cell_styles (id)
            )
        ''')
        # === КОНЕЦ НОВЫХ ТАБЛИЦ ДЛЯ СТИЛЕЙ ===

        # === НОВАЯ ТАБЛИЦА ДЛЯ ОБЪЕДИНЕННЫХ ЯЧЕЕК ===
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS merged_cells_ranges (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet_id INTEGER NOT NULL,
                range_address TEXT NOT NULL, -- Адрес объединенного диапазона (например, 'A1:C3')
                FOREIGN KEY (sheet_id) REFERENCES sheets (id)
            )
        ''')
        # === КОНЕЦ НОВОЙ ТАБЛИЦЫ ===

        self.connection.commit()
        logger.info("Схема базы данных инициализирована успешно.")

    # - НОВЫЕ МЕТОДЫ ДЛЯ РАБОТЫ С СЫРЫМИ ДАННЫМИ -
    def _get_raw_data_table_name(self, sheet_name: str) -> str:
        """Генерирует имя таблицы для сырых данных листа."""
        # Используем санитизированное имя листа как основу
        base_name = sanitize_table_name(sheet_name)
        return f"raw_data_{base_name}"

    def create_raw_data_table(self, sheet_name: str, column_names: List[str]) -> bool:
        """
        Создает таблицу в БД для хранения сырых данных листа.
        Args:
            sheet_name (str): Имя листа Excel.
            column_names (List[str]): Список имен столбцов.
        Returns:
            bool: True, если таблица создана или уже существует, False в случае ошибки.
        """
        if not self.connection:
            logger.error("Нет активного соединения с БД для создания таблицы сырых данных.")
            return False

        try:
            cursor = self.connection.cursor()
            table_name = self._get_raw_data_table_name(sheet_name)
            logger.debug(f"[DEBUG_STORAGE] Создание таблицы сырых данных '{table_name}' для листа '{sheet_name}'.")

            # 1. Проверяем, существует ли таблица
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name=?
            """, (table_name,))
            if cursor.fetchone():
                logger.debug(f"[DEBUG_STORAGE] Таблица '{table_name}' уже существует.")
                return True # Таблица уже существует

            # 2. Создаем таблицу
            # Начинаем с обязательных служебных столбцов
            columns_sql_parts = ["id INTEGER PRIMARY KEY AUTOINCREMENT"] # Уникальный ID строки
            
            # Добавляем столбцы для данных
            for col_name in column_names:
                # Санитизируем имя столбца
                sanitized_col_name = sanitize_table_name(col_name)
                
                # - ИСПРАВЛЕНО: Проверка на конфликт имён -
                # Проверяем, не совпадает ли санитизированное имя с зарезервированными
                if sanitized_col_name.lower() in ['id']:
                    # Если совпадает, добавляем префикс
                    sanitized_col_name = f"data_{sanitized_col_name}"
                    logger.debug(f"[DEBUG_STORAGE] Зарезервированное имя столбца '{col_name}' переименовано в '{sanitized_col_name}' для таблицы '{table_name}'.")
                # - КОНЕЦ ИСПРАВЛЕНИЯ -
                
                # Добавляем столбец (в SQLite все значения TEXT, можно уточнить тип позже)
                columns_sql_parts.append(f"{sanitized_col_name} TEXT")

            create_table_sql = f"CREATE TABLE {table_name} ({', '.join(columns_sql_parts)})"
            logger.debug(f"[DEBUG_STORAGE] SQL-запрос создания таблицы: {create_table_sql}")
            cursor.execute(create_table_sql)
            
            # 3. Регистрируем таблицу в реестре
            # Сначала нужно получить sheet_id. Предполагаем, что лист уже создан.
            # Это может потребовать передачи sheet_id в функцию или поиска по имени.
            # Для простоты, пока оставим регистрацию на усмотрение вызывающего кода,
            # или сделаем её в save_sheet_raw_data.
            
            self.connection.commit()
            logger.info(f"[DEBUG_STORAGE] Таблица сырых данных '{table_name}' успешно создана.")
            return True

        except sqlite3.Error as e:
            logger.error(f"[DEBUG_STORAGE] Ошибка SQLite при создании таблицы '{table_name}': {e}")
            self.connection.rollback()
            return False
        except Exception as e:
            logger.error(f"[DEBUG_STORAGE] Неожиданная ошибка при создании таблицы '{table_name}': {e}")
            self.connection.rollback()
            return False

    def save_sheet_raw_data(self, sheet_id: int, sheet_name: str, raw_data_info: Dict[str, Any]) -> bool:
        """
        Сохраняет сырые данные листа в отдельную таблицу.
        Args:
            sheet_id (int): ID листа в БД.
            sheet_name (str): Имя листа Excel.
            raw_data_info (Dict[str, Any]): Информация о сырых данных (ключи: 'column_names', 'rows').
        Returns:
            bool: True, если данные сохранены успешно, False в случае ошибки.
        """
        if not self.connection:
            logger.error("Нет активного соединения с БД для сохранения сырых данных.")
            return False

        try:
            table_name = self._get_raw_data_table_name(sheet_name)
            column_names = raw_data_info.get("column_names", [])
            rows_data = raw_data_info.get("rows", [])
            
            logger.debug(f"[DEBUG_STORAGE] Начало сохранения сырых данных для листа '{sheet_name}' (ID: {sheet_id}) в таблицу '{table_name}'.")

            # 1. Создаем таблицу (если не существует)
            # Передаем column_names для создания столбцов
            if not self.create_raw_data_table(sheet_name, column_names):
                 logger.error(f"[DEBUG_STORAGE] Не удалось создать таблицу для сырых данных листа '{sheet_name}'.")
                 return False

            # 2. Вставляем данные
            cursor = self.connection.cursor()
            
            # Подготавливаем имена столбцов для вставки (санитизированные)
            sanitized_col_names = []
            for cn in column_names:
                s_cn = sanitize_table_name(cn)
                # Применяем ту же логику переименования, что и при создании
                if s_cn.lower() == 'id': # Только 'id' конфликтует, 'row_index' не создается как столбец
                    s_cn = f"data_{s_cn}"
                sanitized_col_names.append(s_cn)
            
            logger.debug(f"[DEBUG_STORAGE] Санитизированные имена столбцов для вставки: {sanitized_col_names}")

            if not sanitized_col_names:
                logger.warning(f"[DEBUG_STORAGE] Нет санитизированных имен столбцов для вставки в '{table_name}'.")
                return True # Нечего вставлять, но это не ошибка

            # Формируем SQL-запрос для вставки
            # placeholders - это список '?' для VALUES
            placeholders = ', '.join(['?' for _ in sanitized_col_names])
            # columns_str - это список санитизированных имен столбцов
            columns_str = ', '.join(sanitized_col_names)
            
            insert_sql = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
            logger.debug(f"[DEBUG_STORAGE] SQL-запрос вставки данных: {insert_sql}")

            # Подготавливаем данные для вставки
            data_to_insert = []
            for row_dict in rows_data:
                # Для каждой строки создаем кортеж значений в порядке sanitized_col_names
                row_values = []
                for orig_col_name in column_names: # Итерируемся по оригинальным именам
                    sanitized_name = sanitize_table_name(orig_col_name)
                    if sanitized_name.lower() == 'id':
                        sanitized_name = f"data_{sanitized_name}"
                    
                    # Получаем значение из row_dict по оригинальному имени
                    value = row_dict.get(orig_col_name, None)
                    # Убедимся, что значение является строкой или None для SQLite
                    if value is not None and not isinstance(value, (str, int, float, type(None))):
                         # Если это datetime, преобразуем в ISO строку
                         if isinstance(value, datetime):
                              value = value.isoformat()
                         else:
                              # Для остальных типов - преобразуем в строку
                              value = str(value)
                    row_values.append(value)
                data_to_insert.append(tuple(row_values))
            
            logger.debug(f"[DEBUG_STORAGE] Подготовлено {len(data_to_insert)} строк для вставки.")

            # Выполняем массовую вставку
            if data_to_insert: # Проверяем, есть ли данные для вставки
                cursor.executemany(insert_sql, data_to_insert)
                self.connection.commit()
                logger.info(f"[DEBUG_STORAGE] В таблицу '{table_name}' вставлено {len(data_to_insert)} строк сырых данных.")
            else:
                 logger.info(f"[DEBUG_STORAGE] Нет данных для вставки в таблицу '{table_name}'.")
            
            # 3. Регистрируем таблицу в реестре (если ещё не зарегистрирована)
            cursor.execute("""
                INSERT OR IGNORE INTO raw_data_tables_registry (sheet_id, table_name) 
                VALUES (?, ?)
            """, (sheet_id, table_name))
            self.connection.commit()
            
            logger.info(f"[DEBUG_STORAGE] Сырые данные для листа '{sheet_name}' (ID: {sheet_id}) успешно сохранены в таблицу '{table_name}'.")
            return True

        except sqlite3.Error as e:
            logger.error(f"[DEBUG_STORAGE] Ошибка SQLite при сохранении сырых данных для листа '{sheet_name}' (ID: {sheet_id}): {e}")
            self.connection.rollback()
            return False
        except Exception as e:
            logger.error(f"[DEBUG_STORAGE] Неожиданная ошибка при сохранении сырых данных для листа '{sheet_name}' (ID: {sheet_id}): {e}")
            self.connection.rollback()
            return False

    def load_sheet_raw_data(self, sheet_name: str) -> Dict[str, Any]:
        """
        Загружает сырые данные листа из его таблицы.
        """
        raw_data_info = {"column_names": [], "rows": []}
        if not self.connection:
            logger.error("Нет активного соединения с БД для загрузки сырых данных.")
            return raw_data_info

        try:
            table_name = self._get_raw_data_table_name(sheet_name)
            logger.debug(f"[DEBUG_STORAGE] Начало загрузки сырых данных для листа '{sheet_name}' из таблицы '{table_name}'.")

            # Проверяем, существует ли таблица
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name=?
            """, (table_name,))
            if not cursor.fetchone():
                logger.warning(f"[DEBUG_STORAGE] Таблица сырых данных '{table_name}' не найдена.")
                return raw_data_info # Возвращаем пустую структуру

            # Получаем имена столбцов
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns_info = cursor.fetchall()
            # columns_info - список кортежей (cid, name, type, notnull, dflt_value, pk)
            # Исключаем служебный столбец 'id'
            column_names = [col_info[1] for col_info in columns_info if col_info[1].lower() != 'id']
            # Убираем префикс 'data_' если он был добавлен
            original_column_names = [cn[5:] if cn.startswith('data_') else cn for cn in column_names]
            
            raw_data_info["column_names"] = original_column_names
            logger.debug(f"[DEBUG_STORAGE] Загружены имена столбцов: {original_column_names}")

            # Получаем все строки данных
            # Формируем список столбцов для SELECT
            select_columns = ', '.join(column_names) if column_names else '*'
            cursor.execute(f"SELECT {select_columns} FROM {table_name}")
            rows = cursor.fetchall()
            
            # Преобразуем кортежи в словари
            rows_data = []
            for row_tuple in rows:
                row_dict = {}
                for i, col_name in enumerate(column_names):
                    orig_col_name = original_column_names[i]
                    value = row_tuple[i]
                    # Здесь можно добавить десериализацию из строки, если потребуется
                    row_dict[orig_col_name] = value
                rows_data.append(row_dict)
                
            raw_data_info["rows"] = rows_data
            logger.info(f"[DEBUG_STORAGE] Сырые данные для листа '{sheet_name}' загружены. Всего строк: {len(raw_data_info['rows'])}")
            return raw_data_info

        except sqlite3.Error as e:
            logger.error(f"[DEBUG_STORAGE] Ошибка SQLite при загрузке сырых данных для листа '{sheet_name}': {e}")
            # Возвращаем пустую структуру в случае ошибки
            return {"column_names": [], "rows": []}
        except Exception as e:
            logger.error(f"[DEBUG_STORAGE] Неожиданная ошибка при загрузке сырых данных для листа '{sheet_name}': {e}")
            # Возвращаем пустую структуру в случае ошибки
            return {"column_names": [], "rows": []}
    # - КОНЕЦ НОВЫХ МЕТОДОВ ДЛЯ СЫРЫХ ДАННЫХ -

    # - НОВЫЕ МЕТОДЫ ДЛЯ РАБОТЫ СО СТИЛЯМИ -
    def _get_or_create_font(self, cursor, font_attrs: Dict[str, Any]) -> Optional[int]:
        """Получает ID существующего шрифта или создает новый."""
        if not font_attrs:
            return None
        columns = list(font_attrs.keys())
        placeholders = ', '.join(['?' for _ in columns])
        select_conditions = ' AND '.join([f"{col} IS ?" for col in columns]) # IS для корректной проверки NULL
        insert_sql = f"INSERT OR IGNORE INTO fonts ({', '.join(columns)}) VALUES ({placeholders})"
        select_sql = f"SELECT id FROM fonts WHERE {select_conditions}"
        
        values = [font_attrs.get(col) for col in columns]
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values)
        row = cursor.fetchone()
        return row[0] if row else None

    def _get_or_create_fill(self, cursor, fill_attrs: Dict[str, Any]) -> Optional[int]:
        """Получает ID существующей заливки или создает новую."""
        if not fill_attrs:
            return None
        columns = list(fill_attrs.keys())
        placeholders = ', '.join(['?' for _ in columns])
        select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
        insert_sql = f"INSERT OR IGNORE INTO pattern_fills ({', '.join(columns)}) VALUES ({placeholders})"
        select_sql = f"SELECT id FROM pattern_fills WHERE {select_conditions}"
        
        values = [fill_attrs.get(col) for col in columns]
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values)
        row = cursor.fetchone()
        return row[0] if row else None

    def _get_or_create_border(self, cursor, border_attrs: Dict[str, Any]) -> Optional[int]:
        """Получает ID существующих границ или создает новые."""
        if not border_attrs:
            return None
        columns = list(border_attrs.keys())
        placeholders = ', '.join(['?' for _ in columns])
        select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
        insert_sql = f"INSERT OR IGNORE INTO borders ({', '.join(columns)}) VALUES ({placeholders})"
        select_sql = f"SELECT id FROM borders WHERE {select_conditions}"
        
        values = [border_attrs.get(col) for col in columns]
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values)
        row = cursor.fetchone()
        return row[0] if row else None

    def _get_or_create_alignment(self, cursor, align_attrs: Dict[str, Any]) -> Optional[int]:
        """Получает ID существующего выравнивания или создает новое."""
        if not align_attrs:
            return None
        columns = list(align_attrs.keys())
        placeholders = ', '.join(['?' for _ in columns])
        select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
        insert_sql = f"INSERT OR IGNORE INTO alignments ({', '.join(columns)}) VALUES ({placeholders})"
        select_sql = f"SELECT id FROM alignments WHERE {select_conditions}"
        
        values = [align_attrs.get(col) for col in columns]
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values)
        row = cursor.fetchone()
        return row[0] if row else None

    def _get_or_create_protection(self, cursor, prot_attrs: Dict[str, Any]) -> Optional[int]:
        """Получает ID существующей защиты или создает новую."""
        if not prot_attrs:
            return None
        # Приводим BOOLEAN значения к INTEGER
        prot_attrs = {k: int(v) if isinstance(v, bool) else v for k, v in prot_attrs.items()}
        columns = list(prot_attrs.keys())
        placeholders = ', '.join(['?' for _ in columns])
        select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
        insert_sql = f"INSERT OR IGNORE INTO protections ({', '.join(columns)}) VALUES ({placeholders})"
        select_sql = f"SELECT id FROM protections WHERE {select_conditions}"
        
        values = [prot_attrs.get(col) for col in columns]
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values)
        row = cursor.fetchone()
        return row[0] if row else None

    def _get_or_create_cell_style(self, cursor, style_attrs: Dict[str, Any]) -> Optional[int]:
        """Получает ID существующего стиля ячейки или создает новый."""
        if not style_attrs:
            return None
            
        # Извлекаем атрибуты для каждого компонента стиля
        font_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('font_')}
        fill_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('fill_')}
        border_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('border_')}
        align_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('alignment_')}
        prot_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('protection_')}
        
        # Получаем или создаем ID для каждого компонента
        font_id = self._get_or_create_font(cursor, font_attrs)
        fill_id = self._get_or_create_fill(cursor, fill_attrs)
        border_id = self._get_or_create_border(cursor, border_attrs)
        alignment_id = self._get_or_create_alignment(cursor, align_attrs)
        protection_id = self._get_or_create_protection(cursor, prot_attrs)
        
        # Атрибуты самого стиля (те, что не входят в подкомпоненты)
        style_main_attrs = {
            "font_id": font_id,
            "fill_id": fill_id,
            "border_id": border_id,
            "alignment_id": alignment_id,
            "protection_id": protection_id,
            "num_fmt_id": style_attrs.get("num_fmt_id"),
            "xf_id": style_attrs.get("xf_id"),
            "quote_prefix": int(style_attrs.get("quote_prefix", 0)) if style_attrs.get("quote_prefix") is not None else None,
        }
        
        columns = list(style_main_attrs.keys())
        placeholders = ', '.join(['?' for _ in columns])
        select_conditions = ' AND '.join([f"{col} IS ?" for col in columns])
        insert_sql = f"INSERT OR IGNORE INTO cell_styles ({', '.join(columns)}) VALUES ({placeholders})"
        select_sql = f"SELECT id FROM cell_styles WHERE {select_conditions}"
        
        values = [style_main_attrs.get(col) for col in columns]
        cursor.execute(insert_sql, values)
        cursor.execute(select_sql, values)
        row = cursor.fetchone()
        return row[0] if row else None

    def save_sheet_styles(self, sheet_id: int, styled_ranges_data: List[Dict[str, Any]]) -> bool:
        """
        Сохраняет уникальные стили и их применение к диапазонам на листе.
        Args:
            sheet_id (int): ID листа в БД.
            styled_ranges_data (List[Dict[str, Any]]): Список словарей, где каждый словарь
                содержит 'style_attributes' (dict) и 'range_address' (str).
        Returns:
            bool: True, если данные сохранены успешно, False в случае ошибки.
        """
        if not self.connection:
            logger.error("Нет активного соединения с БД для сохранения стилей.")
            return False

        try:
            cursor = self.connection.cursor()
            for style_range_info in styled_ranges_data:
                style_attrs = style_range_info.get("style_attributes", {})
                range_addr = style_range_info.get("range_address", "")

                if not style_attrs or not range_addr:
                    logger.warning(f"Пропущен стиль/диапазон: {style_range_info}")
                    continue

                # 1. Получаем или создаем ID стиля
                style_id = self._get_or_create_cell_style(cursor, style_attrs)
                if style_id is None:
                    logger.error(f"Не удалось получить или создать стиль для: {style_attrs}")
                    continue

                # 2. Сохраняем связь стиль-диапазон
                cursor.execute('''
                    INSERT OR IGNORE INTO styled_ranges (sheet_id, style_id, range_address) 
                    VALUES (?, ?, ?)
                ''', (sheet_id, style_id, range_addr))

            self.connection.commit()
            logger.info(f"Стили для листа ID {sheet_id} успешно сохранены.")
            return True

        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при сохранении стилей для листа ID {sheet_id}: {e}")
            self.connection.rollback()
            return False
        except Exception as e:
            logger.error(f"Неожиданная ошибка при сохранении стилей для листа ID {sheet_id}: {e}")
            self.connection.rollback()
            return False

    def load_sheet_styles(self, sheet_id: int) -> List[Dict[str, Any]]:
        """
        Загружает стили и диапазоны для указанного листа.
        Args:
            sheet_id (int): ID листа в БД.
        Returns:
            List[Dict[str, Any]]: Список словарей с 'style_attributes' и 'range_address'.
        """
        styles_data = []
        if not self.connection:
            logger.error("Нет активного соединения с БД для загрузки стилей.")
            return styles_data

        try:
            cursor = self.connection.cursor()
            
            # Запрашиваем связанные стили и диапазоны
            # Это сложный JOIN, который собирает всю информацию о стиле
            cursor.execute('''
                SELECT 
                    sr.range_address,
                    f.name, f.sz, f.b, f.i, f.u, f.strike, f.color, f.color_theme, f.color_tint, f.vert_align, f.scheme,
                    pf.pattern_type, pf.fg_color, pf.fg_color_theme, pf.fg_color_tint, pf.bg_color, pf.bg_color_theme, pf.bg_color_tint,
                    b.left_style, b.left_color, b.right_style, b.right_color, b.top_style, b.top_color,
                    b.bottom_style, b.bottom_color, b.diagonal_style, b.diagonal_color, b.diagonal_up, b.diagonal_down, b.outline,
                    a.horizontal, a.vertical, a.text_rotation, a.wrap_text, a.shrink_to_fit, a.indent,
                    a.relative_indent, a.justify_last_line, a.reading_order, a.text_direction,
                    p.locked, p.hidden,
                    cs.num_fmt_id, cs.xf_id, cs.quote_prefix
                FROM styled_ranges sr
                LEFT JOIN cell_styles cs ON sr.style_id = cs.id
                LEFT JOIN fonts f ON cs.font_id = f.id
                LEFT JOIN pattern_fills pf ON cs.fill_id = pf.id
                LEFT JOIN borders b ON cs.border_id = b.id
                LEFT JOIN alignments a ON cs.alignment_id = a.id
                LEFT JOIN protections p ON cs.protection_id = p.id
                WHERE sr.sheet_id = ?
            ''', (sheet_id,))
            
            rows = cursor.fetchall()
            for row in rows:
                range_addr = row[0]
                
                # Собираем атрибуты стиля из результата запроса
                style_attrs = {}
                # Font (индексы 1-11)
                font_keys = ["name", "sz", "b", "i", "u", "strike", "color", "color_theme", "color_tint", "vert_align", "scheme"]
                for i, key in enumerate(font_keys):
                    if row[i+1] is not None: # +1 потому что range_address это 0
                        style_attrs[f"font_{key}"] = row[i+1]
                
                # Fill (индексы 12-18)
                fill_keys = ["pattern_type", "fg_color", "fg_color_theme", "fg_color_tint", "bg_color", "bg_color_theme", "bg_color_tint"]
                for i, key in enumerate(fill_keys):
                    if row[i+12] is not None:
                        style_attrs[f"fill_{key}"] = row[i+12]
                
                # Border (индексы 19-30)
                border_keys = ["left_style", "left_color", "right_style", "right_color", "top_style", "top_color",
                               "bottom_style", "bottom_color", "diagonal_style", "diagonal_color", "diagonal_up", "diagonal_down", "outline"]
                for i, key in enumerate(border_keys):
                    if row[i+19] is not None:
                        style_attrs[f"border_{key}"] = row[i+19]
                
                # Alignment (индексы 31-41)
                align_keys = ["horizontal", "vertical", "text_rotation", "wrap_text", "shrink_to_fit", "indent",
                              "relative_indent", "justify_last_line", "reading_order", "text_direction"]
                for i, key in enumerate(align_keys):
                    if row[i+31] is not None:
                        style_attrs[f"alignment_{key}"] = row[i+31]
                
                # Protection (индексы 42-43)
                prot_keys = ["locked", "hidden"]
                for i, key in enumerate(prot_keys):
                    if row[i+42] is not None:
                        style_attrs[f"protection_{key}"] = row[i+42]
                
                # Cell Style main attrs (индексы 44-46)
                cs_keys = ["num_fmt_id", "xf_id", "quote_prefix"]
                for i, key in enumerate(cs_keys):
                    if row[i+44] is not None:
                        style_attrs[key] = row[i+44]
                
                styles_data.append({
                    "style_attributes": style_attrs,
                    "range_address": range_addr
                })
            
            logger.info(f"Загружено {len(styles_data)} стилей для листа ID {sheet_id}.")
            return styles_data

        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при загрузке стилей для листа ID {sheet_id}: {e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка при загрузке стилей для листа ID {sheet_id}: {e}")
            return []
    # - КОНЕЦ НОВЫХ МЕТОДОВ ДЛЯ СТИЛЕЙ -

    # - НОВЫЕ МЕТОДЫ ДЛЯ РАБОТЫ С ОБЪЕДИНЕННЫМИ ЯЧЕЙКАМИ -
    def save_sheet_merged_cells(self, sheet_id: int, merged_ranges: List[str]) -> bool:
        """
        Сохраняет диапазоны объединенных ячеек для листа.
        Args:
            sheet_id (int): ID листа в БД.
            merged_ranges (List[str]): Список строковых представлений диапазонов.
        Returns:
            bool: True, если данные сохранены успешно, False в случае ошибки.
        """
        if not self.connection:
            logger.error("Нет активного соединения с БД для сохранения объединенных ячеек.")
            return False

        try:
            cursor = self.connection.cursor()
            for range_addr in merged_ranges:
                if not range_addr:
                    continue
                cursor.execute('''
                    INSERT OR IGNORE INTO merged_cells_ranges (sheet_id, range_address) 
                    VALUES (?, ?)
                ''', (sheet_id, range_addr))

            self.connection.commit()
            logger.info(f"Объединенные ячейки для листа ID {sheet_id} успешно сохранены.")
            return True

        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при сохранении объединенных ячеек для листа ID {sheet_id}: {e}")
            self.connection.rollback()
            return False
        except Exception as e:
            logger.error(f"Неожиданная ошибка при сохранении объединенных ячеек для листа ID {sheet_id}: {e}")
            self.connection.rollback()
            return False

    def load_sheet_merged_cells(self, sheet_id: int) -> List[str]:
        """
        Загружает диапазоны объединенных ячеек для указанного листа.
        Args:
            sheet_id (int): ID листа в БД.
        Returns:
            List[str]: Список строковых представлений диапазонов.
        """
        merged_ranges = []
        if not self.connection:
            logger.error("Нет активного соединения с БД для загрузки объединенных ячеек.")
            return merged_ranges

        try:
            cursor = self.connection.cursor()
            cursor.execute('SELECT range_address FROM merged_cells_ranges WHERE sheet_id = ?', (sheet_id,))
            rows = cursor.fetchall()
            merged_ranges = [row[0] for row in rows]
            logger.info(f"Загружено {len(merged_ranges)} объединенных диапазонов для листа ID {sheet_id}.")
            return merged_ranges

        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при загрузке объединенных ячеек для листа ID {sheet_id}: {e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка при загрузке объединенных ячеек для листа ID {sheet_id}: {e}")
            return []
    # - КОНЕЦ НОВЫХ МЕТОДОВ ДЛЯ ОБЪЕДИНЕННЫХ ЯЧЕЕК -

    def save_analysis_results(self, project_name: str, documentation_data: Dict[str, Any]):
        """
        Сохраняет результаты анализа документации в базу данных.
        Args:
            project_name (str): Имя проекта.
            documentation_data (Dict[str, Any]): Данные документации, полученные от анализатора.
        """
        if not self.connection:
            logger.error("Нет активного соединения с БД для сохранения результатов.")
            return False # Добавлен возврат значения

        try:
            cursor = self.connection.cursor()

            # 1. Создание/получение записи проекта
            cursor.execute("SELECT id FROM projects WHERE name = ?", (project_name,))
            project_row = cursor.fetchone()
            if project_row:
                project_id = project_row[0]
                logger.debug(f"Проект '{project_name}' уже существует (ID: {project_id}).")
            else:
                # Создаем новый проект
                created_at_iso = datetime.now().isoformat()
                cursor.execute(
                    "INSERT INTO projects (name, created_at) VALUES (?, ?)",
                    (project_name, created_at_iso)
                )
                project_id = cursor.lastrowid
                logger.info(f"Создан новый проект '{project_name}' (ID: {project_id}).")

            # 2. Сохранение информации о листах и связанной информации
            sheets_data = documentation_data.get("sheets", {})
            # Итерируемся по словарю sheets_data
            # sheet_name - ключ, sheet_info - значение (словарь с данными листа)
            for sheet_name, sheet_info in sheets_data.items(): 
                logger.debug(f"Обработка листа: {sheet_name}")

                # 2.1. Создание/получение записи листа (ИЗМЕНЕНО: index -> sheet_index)
                sheet_index = sheet_info.get("index", 0) # Получаем индекс из данных анализатора
                cursor.execute(
                    "SELECT id FROM sheets WHERE project_id = ? AND name = ?",
                    (project_id, sheet_name)
                )
                sheet_row = cursor.fetchone()
                if sheet_row:
                    sheet_id = sheet_row[0]
                    logger.debug(f"  Лист '{sheet_name}' уже существует (ID: {sheet_id}).")
                     # Опционально: обновить информацию о листе, если она изменилась
                    cursor.execute(
                        "UPDATE sheets SET sheet_index = ? WHERE id = ?",
                        (sheet_index, sheet_id) # Обновляем индекс, если он важен
                    )
                else:
                    # Создаем новую запись листа
                    # Сериализуем структуру в JSON, используя наш кастомный энкодер
                    structure_json = json.dumps(sheet_info.get("structure", []), cls=DateTimeEncoder, ensure_ascii=False)
                    # НОВОЕ: Сериализуем информацию о сырых данных (только имена столбцов и общее кол-во)
                    raw_data_summary = {
                        "column_names": sheet_info.get("raw_data", {}).get("column_names", []),
                        # Можно добавить кол-во строк, если нужно быстро получить сводку
                        # "row_count": len(sheet_info.get("raw_data", {}).get("rows", []))
                    }
                    raw_data_info_json = json.dumps(raw_data_summary, cls=DateTimeEncoder, ensure_ascii=False)
                    
                    cursor.execute(
                        "INSERT INTO sheets (project_id, name, sheet_index, structure, raw_data_info) VALUES (?, ?, ?, ?, ?)",
                        (project_id, sheet_name, sheet_index, structure_json, raw_data_info_json) # ИЗМЕНЕНО
                    )
                    sheet_id = cursor.lastrowid
                    logger.debug(f"  Создан новый лист '{sheet_name}' (ID: {sheet_id}).")

                # 2.2. Сохранение формул листа
                formulas_data = sheet_info.get("formulas", [])
                # Сначала удаляем старые формулы для этого листа (операция UPSERT)
                cursor.execute('DELETE FROM formulas WHERE sheet_id = ?', (sheet_id,))
                logger.debug(f"  Удалены старые формулы для листа '{sheet_name}' (ID: {sheet_id}).")

                for formula_info in formulas_data:
                    cell = formula_info.get("cell", "")
                    formula = formula_info.get("formula", "")
                    references = formula_info.get("references", [])
                    # Сериализуем ссылки в JSON
                    references_json = json.dumps(references, cls=DateTimeEncoder, ensure_ascii=False)
                    # === ИСПРАВЛЕНО: Экранирование имени столбца ===
                    cursor.execute(
                        'INSERT INTO formulas (sheet_id, cell, formula, "references") VALUES (?, ?, ?, ?)', # <-- "references" в кавычках
                        (sheet_id, cell, formula, references_json)
                    )
                    # =================================================

                logger.debug(f"  Сохранено {len(formulas_data)} формул для листа '{sheet_name}'.")

                # 2.3. Сохранение межлистовых ссылок
                cross_refs_data = sheet_info.get("cross_sheet_references", [])
                # Сначала удаляем старые ссылки для этого листа
                cursor.execute("DELETE FROM cross_sheet_references WHERE sheet_id = ?", (sheet_id,))
                logger.debug(f"  Удалены старые межлистовые ссылки для листа '{sheet_name}'.")

                for ref_info in cross_refs_data:
                    from_cell = ref_info.get("from_cell", "")
                    from_formula = ref_info.get("from_formula", "")
                    to_sheet = ref_info.get("to_sheet", "")
                    ref_type = ref_info.get("reference_type", "")
                    ref_address = ref_info.get("reference_address", "")
                    cursor.execute(
                        "INSERT INTO cross_sheet_references (sheet_id, from_cell, from_formula, to_sheet, reference_type, reference_address) VALUES (?, ?, ?, ?, ?, ?)",
                        (sheet_id, from_cell, from_formula, to_sheet, ref_type, ref_address)
                    )
                logger.debug(f"  Сохранено {len(cross_refs_data)} межлистовых ссылок для листа '{sheet_name}'.")

                # === ИЗМЕНЕНО: Сохранение диаграмм в структурированном виде ===
                charts_data = sheet_info.get("charts", [])
                # Сначала удаляем старые диаграммы и их источники для этого листа
                # Удаляем источники данных (из-за внешнего ключа)
                chart_ids_to_delete = []
                cursor.execute("SELECT id FROM charts WHERE sheet_id = ?", (sheet_id,))
                for row in cursor.fetchall():
                    chart_ids_to_delete.append(row[0])
                if chart_ids_to_delete:
                    placeholders = ','.join('?' * len(chart_ids_to_delete))
                    # Удаляем сначала серии, потом источники, потом диаграммы
                    cursor.execute(f"SELECT id FROM chart_series WHERE chart_id IN ({placeholders})", chart_ids_to_delete)
                    series_ids_to_delete = [r[0] for r in cursor.fetchall()]
                    if series_ids_to_delete:
                        series_placeholders = ','.join('?' * len(series_ids_to_delete))
                        cursor.execute(f"DELETE FROM chart_data_sources WHERE series_id IN ({series_placeholders})", series_ids_to_delete)
                    cursor.execute(f"DELETE FROM chart_series WHERE chart_id IN ({placeholders})", chart_ids_to_delete)
                    cursor.execute(f"DELETE FROM chart_axes WHERE chart_id IN ({placeholders})", chart_ids_to_delete)
                    cursor.execute(f"DELETE FROM charts WHERE sheet_id = ?", (sheet_id,))
                
                logger.debug(f"  Удалены старые диаграммы для листа '{sheet_name}'.")

                for chart_info in charts_data:
                    # Извлекаем основные атрибуты диаграммы
                    chart_type = chart_info.get("type", "")
                    chart_title = chart_info.get("title", "")
                    top_left_cell = chart_info.get("top_left_cell", "")
                    width = chart_info.get("width")
                    height = chart_info.get("height")
                    chart_style = chart_info.get("style")
                    legend_position = chart_info.get("legend_position")
                    auto_scaling = chart_info.get("auto_scaling")
                    plot_vis_only = chart_info.get("plot_vis_only")

                    if not chart_type:
                        logger.warning(f"Пропущена диаграмма без типа: {chart_info}")
                        continue

                    # Вставляем запись о диаграмме
                    cursor.execute('''
                        INSERT INTO charts (sheet_id, type, title, top_left_cell, width, height, style, legend_position, auto_scaling, plot_vis_only)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (sheet_id, chart_type, chart_title, top_left_cell, width, height, chart_style, legend_position, auto_scaling, plot_vis_only))
                    
                    chart_id = cursor.lastrowid
                    logger.debug(f"  Создана запись диаграммы ID {chart_id} типа {chart_type}")

                    # Вставляем оси диаграммы
                    axes_data = chart_info.get("axes", [])
                    for axis_info in axes_data:
                        axis_type = axis_info.get("axis_type")
                        ax_id = axis_info.get("ax_id")
                        ax_pos = axis_info.get("ax_pos")
                        delete_axis = axis_info.get("delete")
                        axis_title = axis_info.get("title")
                        num_fmt = axis_info.get("num_fmt")
                        major_tick_mark = axis_info.get("major_tick_mark")
                        minor_tick_mark = axis_info.get("minor_tick_mark")
                        tick_lbl_pos = axis_info.get("tick_lbl_pos")
                        crosses = axis_info.get("crosses")
                        crosses_at = axis_info.get("crosses_at")
                        major_unit = axis_info.get("major_unit")
                        minor_unit = axis_info.get("minor_unit")
                        min_val = axis_info.get("min")
                        max_val = axis_info.get("max")
                        orientation = axis_info.get("orientation")
                        log_base = axis_info.get("log_base")
                        major_gridlines = axis_info.get("major_gridlines")
                        minor_gridlines = axis_info.get("minor_gridlines")
                        
                        cursor.execute('''
                            INSERT INTO chart_axes (
                                chart_id, axis_type, ax_id, ax_pos, delete_axis, title, num_fmt,
                                major_tick_mark, minor_tick_mark, tick_lbl_pos, crosses, crosses_at,
                                major_unit, minor_unit, min, max, orientation, log_base,
                                major_gridlines, minor_gridlines
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            chart_id, axis_type, ax_id, ax_pos, delete_axis, axis_title, num_fmt,
                            major_tick_mark, minor_tick_mark, tick_lbl_pos, crosses, crosses_at,
                            major_unit, minor_unit, min_val, max_val, orientation, log_base,
                            major_gridlines, minor_gridlines
                        ))
                        logger.debug(f"    Добавлена ось {axis_type} для диаграммы ID {chart_id}")

                    # Вставляем серии данных для этой диаграммы
                    series_data = chart_info.get("series", [])
                    data_sources_data = chart_info.get("data_sources", [])
                    for series_info in series_data:
                        series_idx = series_info.get("idx")
                        series_order = series_info.get("order")
                        series_tx = series_info.get("tx")
                        shape = series_info.get("shape")
                        smooth = series_info.get("smooth")
                        invert_if_negative = series_info.get("invert_if_negative")

                        cursor.execute('''
                            INSERT INTO chart_series (chart_id, idx, "order", tx, shape, smooth, invert_if_negative)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        ''', (chart_id, series_idx, series_order, series_tx, shape, smooth, invert_if_negative))
                        
                        series_id = cursor.lastrowid
                        logger.debug(f"    Добавлена серия ID {series_id} (idx {series_idx}) для диаграммы ID {chart_id}")

                        # Вставляем источники данных, связанные с этой серией
                        for ds_info in data_sources_data:
                            if ds_info.get("series_index") == series_idx: # Связываем по индексу серии
                                data_type = ds_info.get("data_type")
                                formula = ds_info.get("formula")
                                cursor.execute('''
                                    INSERT INTO chart_data_sources (series_id, data_type, formula)
                                    VALUES (?, ?, ?)
                                ''', (series_id, data_type, formula))
                                logger.debug(f"      Добавлен источник данных ({data_type}) для серии ID {series_id}")

                logger.debug(f"  Сохранено {len(charts_data)} диаграмм для листа '{sheet_name}'.")
                # === КОНЕЦ ИЗМЕНЕНИЙ ДЛЯ ДИАГРАММ ===
                
                # === ИЗМЕНЕНО: Сохранение сырых данных -
                raw_data_info_to_save = sheet_info.get("raw_data", {})
                if raw_data_info_to_save:
                     # ==========================================
                     # ==== ИСПРАВЛЕНИЕ ОШИБКИ ЗДЕСЬ ============
                     # ==========================================
                     success = self.save_sheet_raw_data(sheet_id, sheet_name, raw_data_info_to_save)
                     # ==========================================
                     # ==========================================
                     if success:
                          logger.debug(f"  Сырые данные для листа '{sheet_name}' сохранены.")
                     else:
                          logger.error(f"  Ошибка при сохранении сырых данных для листа '{sheet_name}'.")
                else:
                     logger.debug(f"  Нет сырых данных для сохранения для листа '{sheet_name}'.")
                # === КОНЕЦ ИЗМЕНЕНИЙ ДЛЯ СЫРЫХ ДАННЫХ ===

                # === НОВОЕ: Сохранение стилей ===
                styled_ranges_data = sheet_info.get("styled_ranges", [])
                if styled_ranges_data:
                    success = self.save_sheet_styles(sheet_id, styled_ranges_data)
                    if success:
                        logger.debug(f"  Стили для листа '{sheet_name}' сохранены.")
                    else:
                        logger.error(f"  Ошибка при сохранении стилей для листа '{sheet_name}'.")
                else:
                    logger.debug(f"  Нет стилей для сохранения для листа '{sheet_name}'.")
                # === КОНЕЦ НОВОГО ДЛЯ СТИЛЕЙ ===

                # === НОВОЕ: Сохранение объединенных ячеек ===
                merged_cells_data = sheet_info.get("merged_cells", [])
                if merged_cells_data:
                    success = self.save_sheet_merged_cells(sheet_id, merged_cells_data)
                    if success:
                        logger.debug(f"  Объединенные ячейки для листа '{sheet_name}' сохранены.")
                    else:
                        logger.error(f"  Ошибка при сохранении объединенных ячеек для листа '{sheet_name}'.")
                else:
                    logger.debug(f"  Нет объединенных ячеек для сохранения для листа '{sheet_name}'.")
                # === КОНЕЦ НОВОГО ДЛЯ ОБЪЕДИНЕННЫХ ЯЧЕЕК ===

            # Подтверждаем транзакцию
            self.connection.commit()
            logger.info("Все результаты анализа успешно сохранены в базу данных.")
            return True # Добавлен возврат значения успешного завершения

        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при сохранении результатов анализа: {e}")
            self.connection.rollback() # Откатываем транзакцию в случае ошибки
            return False # Добавлен возврат значения ошибки
        except Exception as e:
            logger.error(f"Неожиданная ошибка при сохранении результатов анализа: {e}", exc_info=True)
            self.connection.rollback() # Откатываем транзакцию в случае ошибки
            return False # Добавлен возврат значения ошибки

    # === ДОБАВЛЕННЫЙ МЕТОД ===
    def get_all_data(self) -> Dict[str, Any]:
        """
        Загружает все данные проекта из базы данных.
        Returns:
            Dict[str, Any]: Словарь со всеми данными проекта.
        """
        logger.debug("Начало загрузки всех данных проекта из БД.")
        all_data = {
            "project_info": {},
            "sheets": {}
        }
        if not self.connection:
            logger.error("Нет активного соединения с БД для загрузки обзора проекта.")
            return {}

        try:
            cursor = self.connection.cursor()

            # Получаем имя проекта (предполагаем, что в БД только один проект для этого файла)
            cursor.execute("SELECT name FROM projects LIMIT 1")
            project_row = cursor.fetchone()
            if not project_row:
                logger.warning("В базе данных не найдено информации о проекте.")
                return {}
            project_name = project_row[0]
            all_data["project_info"]["name"] = project_name
            logger.debug(f"Загруженное имя проекта: {project_name}")

            # Получаем список имен листов (ИЗМЕНЕНО: index -> sheet_index)
            cursor.execute("SELECT name FROM sheets ORDER BY sheet_index") # ИЗМЕНЕНО
            sheets_rows = cursor.fetchall()
            sheet_names = [row[0] for row in sheets_rows]
            logger.debug(f"Найденные листы: {sheet_names}")

            # Загружаем данные для каждого листа
            for sheet_name in sheet_names:
                logger.debug(f"Загрузка данных для листа: '{sheet_name}'")
                sheet_data = self._load_sheet_data(sheet_name)
                if sheet_data:
                    # Используем имя листа как ключ в словаре all_data["sheets"]
                    all_data["sheets"][sheet_name] = sheet_data 
                else:
                     logger.warning(f"Данные для листа '{sheet_name}' не были загружены или оказались пустыми.")

            logger.debug("Загрузка всех данных проекта завершена.")
            return all_data

        except Exception as e:
            logger.error(f"Ошибка при загрузке всех данных проекта: {e}", exc_info=True)
            return {} # Возвращаем пустой словарь в случае ошибки
    
    def _load_sheet_data(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """
        Загружает данные одного листа из базы данных.
        Args:
            sheet_name (str): Имя листа.
        Returns:
            Optional[Dict[str, Any]]: Словарь с данными листа или None в случае ошибки.
        """
        logger.debug(f"Начало загрузки данных листа '{sheet_name}'")
        sheet_data = {
            "name": sheet_name,
            "structure": [],
            "raw_data": {}, # НОВОЕ
            "formulas": [],
            "cross_sheet_references": [],
            "charts": [], # ИЗМЕНЕНО
            "styled_ranges": [], # НОВОЕ
            "merged_cells": [] # НОВОЕ
        }
        if not self.connection:
            logger.error("Нет активного соединения с БД.")
            return None

        try:
            cursor = self.connection.cursor()

            # Получаем ID листа и другую информацию
            cursor.execute("SELECT id, structure, raw_data_info FROM sheets WHERE name = ?", (sheet_name,)) # ИЗМЕНЕНО
            sheet_row = cursor.fetchone()
            if not sheet_row:
                logger.warning(f"Информация о листе '{sheet_name}' не найдена в БД.")
                return None

            sheet_id, structure_json, raw_data_info_json = sheet_row # ИЗМЕНЕНО
            logger.debug(f"ID листа '{sheet_name}': {sheet_id}")

            # Десериализуем структуру
            if structure_json:
                try:
                    sheet_data["structure"] = json.loads(structure_json)
                    logger.debug(f" Загружена структура листа '{sheet_name}' ({len(sheet_data['structure'])} элементов).")
                except json.JSONDecodeError as e:
                    logger.error(f"Ошибка десериализации структуры листа '{sheet_name}': {e}")
            else:
                 logger.debug(f" Для листа '{sheet_name}' структура отсутствует в БД.")

            # - ИЗМЕНЕНО: Загрузка информации о сырых данных -
            if raw_data_info_json:
                try:
                    raw_data_summary = json.loads(raw_data_info_json)
                    sheet_data["raw_data"]["column_names"] = raw_data_summary.get("column_names", [])
                    # row_count не сохранялся, поэтому его нет
                    logger.debug(f" Загружена сводная информация о сырых данных листа '{sheet_name}'.")
                except json.JSONDecodeError as e:
                    logger.error(f"Ошибка десериализации сводки сырых данных листа '{sheet_name}': {e}")
            else:
                 logger.debug(f" Для листа '{sheet_name}' сводка сырых данных отсутствует в БД.")
            
            # Загружаем сами сырые данные из отдельной таблицы
            loaded_raw_data = self.load_sheet_raw_data(sheet_name)
            if loaded_raw_data and (loaded_raw_data.get("column_names") or loaded_raw_data.get("rows")):
                 sheet_data["raw_data"] = loaded_raw_data
                 logger.debug(f" Загружены полные сырые данные листа '{sheet_name}' ({len(loaded_raw_data.get('rows', []))} строк).")
            else:
                 logger.debug(f" Полные сырые данные для листа '{sheet_name}' отсутствуют или пусты.")
            # - КОНЕЦ ИЗМЕНЕНИЙ -
            
            # Загружаем формулы
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
                sheet_data['formulas'].append({
                    'cell': cell,
                    'formula': formula,
                    'references': references
                })
            logger.debug(f" Загружено {len(sheet_data['formulas'])} формул для листа '{sheet_name}'.")

            # Загружаем межлистовые ссылки
            cursor.execute("SELECT from_cell, from_formula, to_sheet, reference_type, reference_address FROM cross_sheet_references WHERE sheet_id = ?", (sheet_id,))
            cross_refs_rows = cursor.fetchall()
            for row in cross_refs_rows:
                from_cell, from_formula, to_sheet, ref_type, ref_address = row
                sheet_data['cross_sheet_references'].append({
                    'from_cell': from_cell,
                    'from_formula': from_formula,
                    'to_sheet': to_sheet,
                    'reference_type': ref_type,
                    'reference_address': ref_address
                })
            logger.debug(f" Загружено {len(sheet_data['cross_sheet_references'])} межлистовых ссылок для листа '{sheet_name}'.")

            # === ИЗМЕНЕНО: Загрузка диаграмм из структурированных таблиц ===
            # Загружаем основную информацию о диаграммах
            cursor.execute('''
                SELECT id, type, title, top_left_cell, width, height, style, legend_position, auto_scaling, plot_vis_only 
                FROM charts WHERE sheet_id = ?
            ''', (sheet_id,))
            charts_rows = cursor.fetchall()
            
            for chart_row in charts_rows:
                chart_id, chart_type, chart_title, top_left_cell, width, height, chart_style, legend_position, auto_scaling, plot_vis_only = chart_row
                chart_info = {
                    "type": chart_type,
                    "title": chart_title,
                    "top_left_cell": top_left_cell,
                    "width": width,
                    "height": height,
                    "style": chart_style,
                    "legend_position": legend_position,
                    "auto_scaling": auto_scaling,
                    "plot_vis_only": plot_vis_only,
                    "axes": [],
                    "series": [],
                    "data_sources": [] # Будет заполнено позже
                }
                
                # Загружаем оси для этой диаграммы
                cursor.execute('''
                    SELECT axis_type, ax_id, ax_pos, delete_axis, title, num_fmt,
                           major_tick_mark, minor_tick_mark, tick_lbl_pos, crosses, crosses_at,
                           major_unit, minor_unit, min, max, orientation, log_base,
                           major_gridlines, minor_gridlines
                    FROM chart_axes WHERE chart_id = ?
                ''', (chart_id,))
                axes_rows = cursor.fetchall()
                for axis_row in axes_rows:
                    axis_info = {
                        "axis_type": axis_row[0], "ax_id": axis_row[1], "ax_pos": axis_row[2],
                        "delete": axis_row[3], "title": axis_row[4], "num_fmt": axis_row[5],
                        "major_tick_mark": axis_row[6], "minor_tick_mark": axis_row[7],
                        "tick_lbl_pos": axis_row[8], "crosses": axis_row[9], "crosses_at": axis_row[10],
                        "major_unit": axis_row[11], "minor_unit": axis_row[12], "min": axis_row[13],
                        "max": axis_row[14], "orientation": axis_row[15], "log_base": axis_row[16],
                        "major_gridlines": axis_row[17], "minor_gridlines": axis_row[18]
                    }
                    chart_info["axes"].append(axis_info)

                # Загружаем серии данных для этой диаграммы
                cursor.execute('''
                    SELECT id, idx, "order", tx, shape, smooth, invert_if_negative
                    FROM chart_series WHERE chart_id = ? ORDER BY "order"
                ''', (chart_id,))
                series_rows = cursor.fetchall()
                
                series_ids = [] # Собираем ID серий для последующей загрузки источников
                for series_row in series_rows:
                    series_id, series_idx, series_order, series_tx, shape, smooth, invert_if_negative = series_row
                    series_ids.append(series_id)
                    chart_info["series"].append({
                        "idx": series_idx,
                        "order": series_order,
                        "tx": series_tx,
                        "shape": shape,
                        "smooth": smooth,
                        "invert_if_negative": invert_if_negative
                    })
                
                # Загружаем источники данных для всех серий этой диаграммы
                if series_ids:
                    placeholders = ','.join('?' * len(series_ids))
                    cursor.execute(f'''
                        SELECT cds.series_id, cds.data_type, cds.formula, cs.idx
                        FROM chart_data_sources cds
                        JOIN chart_series cs ON cds.series_id = cs.id
                        WHERE cds.series_id IN ({placeholders})
                        ORDER BY cs."order", cds.id
                    ''', series_ids)
                    data_sources_rows = cursor.fetchall()
                    for ds_row in data_sources_rows:
                        series_id, data_type, formula, series_idx = ds_row
                        chart_info["data_sources"].append({
                            "series_index": series_idx, # Используем idx серии для связи
                            "data_type": data_type,
                            "formula": formula
                        })

                sheet_data['charts'].append(chart_info)
            
            logger.debug(f" Загружено {len(sheet_data['charts'])} диаграмм для листа '{sheet_name}'.")
            # === КОНЕЦ ИЗМЕНЕНИЙ ДЛЯ ДИАГРАММ ===

            # === НОВОЕ: Загрузка стилей ===
            loaded_styles = self.load_sheet_styles(sheet_id)
            sheet_data["styled_ranges"] = loaded_styles
            logger.debug(f" Загружено {len(sheet_data['styled_ranges'])} стилей для листа '{sheet_name}'.")
            # === КОНЕЦ НОВОГО ДЛЯ СТИЛЕЙ ===

            # === НОВОЕ: Загрузка объединенных ячеек ===
            loaded_merged_cells = self.load_sheet_merged_cells(sheet_id)
            sheet_data["merged_cells"] = loaded_merged_cells
            logger.debug(f" Загружено {len(sheet_data['merged_cells'])} объединенных диапазонов для листа '{sheet_name}'.")
            # === КОНЕЦ НОВОГО ДЛЯ ОБЪЕДИНЕННЫХ ЯЧЕЕК ===

            logger.debug(f"Данные листа '{sheet_name}' загружены успешно.")
            return sheet_data

        except Exception as e:
            logger.error(f"Ошибка при загрузке данных листа '{sheet_name}': {e}", exc_info=True)
            return None
    # === КОНЕЦ ДОБАВЛЕННОГО МЕТОДА ===

# - ТОЧКА ВХОДА ДЛЯ ТЕСТИРОВАНИЯ -
if __name__ == "__main__":
    # Простой тест подключения и инициализации
    print("--- ТЕСТ ХРАНИЛИЩА ---")
    # Определяем путь к тестовой БД относительно корня проекта
    test_db_path = project_root / "data" / "test_db.sqlite"
    print(f"Путь к тестовой БД: {test_db_path}")
    
    try:
        storage = ProjectDBStorage(str(test_db_path))
        storage.connect()
        print("Подключение к тестовой БД установлено и схема инициализирована.")
        storage.disconnect()
        print("Подключение закрыто.")
        
        # Пытаемся удалить тестовый файл БД
        if test_db_path.exists():
            test_db_path.unlink()
            print("Тестовый файл БД удален.")
            
    except Exception as e:
        print(f"Ошибка при тестировании хранилища: {e}")
