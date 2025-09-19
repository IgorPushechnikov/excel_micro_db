# src/storage/base.py
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
from datetime import datetime  # Импортируем класс datetime
# =========================================

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger
from src.storage.schema import initialize_schema  # Импорт новой функции

# === ИМПОРТЫ НОВЫХ МОДУЛЕЙ ===
# ВАЖНО: Импортируем модули, а не отдельные функции на верхнем уровне,
# чтобы избежать циклических импортов. Функции будем импортировать внутри методов.
from src.storage import raw_data, editable_data, history, formulas, styles, charts, metadata, misc
# =============================

logger = get_logger(__name__)

# === ИСПРАВЛЕНО: Класс DateTimeEncoder ===
class DateTimeEncoder(json.JSONEncoder):
    """Пользовательский JSONEncoder для сериализации datetime объектов."""

    # ИСПРАВЛЕНО: Имя параметра должно быть 'o' для совместимости
    def default(self, o):  # <-- ИЗМЕНЕНО С 'obj' НА 'o'
        # Проверяем, является ли объект экземпляром класса datetime
        # ИСПРАВЛЕНО: Имя параметра
        if isinstance(o, datetime):  # <-- ИЗМЕНЕНО С 'obj' НА 'o'
            # Форматируем дату и время в строку в формате ISO 8601
            # return obj.isoformat() # <-- СТАРОЕ
            return o.isoformat()  # <-- НОВОЕ
        # Для всех остальных типов вызываем метод родителя
        # ИСПРАВЛЕНО: Имя параметра
        return super().default(o)  # <-- ИЗМЕНЕНО С 'obj' НА 'o'
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
        sanitized = f"table_{abs(hash(name))}"  # Создаем уникальное имя на основе хеша
        logger.debug(f"[DEBUG_STORAGE] Имя после санитации было пустым или некорректным. Создано новое имя: '{sanitized}'")

    logger.debug(f"[DEBUG_STORAGE] Санитизированное имя таблицы: '{sanitized}'")
    return sanitized

def sanitize_column_name(name: str) -> str:
    """
    Санитизирует имя для использования в качестве имени столбца SQLite.

    Заменяет недопустимые символы на подчеркивания.

    Args:
        name (str): Исходное имя.

    Returns:
        str: Санитизированное имя.
    """
    logger.debug(f"[DEBUG_STORAGE] Санитизация имени столбца: '{name}'")

    if not name:
        logger.warning("[DEBUG_STORAGE] Получено пустое имя для санитации столбца. Возвращаю '_empty'.")
        return "_empty"

    # Заменяем все недопустимые символы (все, кроме букв, цифр и подчеркиваний) на '_'
    sanitized = "".join(ch if ch.isalnum() or ch == '_' else '_' for ch in name)

    # Убедимся, что имя не пустое и не состоит только из подчеркиваний
    if not sanitized or all(c == '_' for c in sanitized):
        sanitized = f"column_{abs(hash(name))}"  # Создаем уникальное имя на основе хеша
        logger.debug(f"[DEBUG_STORAGE] Имя после санитации было пустым или некорректным. Создано новое имя: '{sanitized}'")

    logger.debug(f"[DEBUG_STORAGE] Санитизированное имя столбца: '{sanitized}'")
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
        если они еще не существуют. Вызывает функцию из schema.py.
        """
        if not self.connection:
            logger.error("Нет активного соединения с БД для инициализации схемы.")
            return

        cursor = self.connection.cursor()
        # === ИЗМЕНЕНО: Вызов функции из schema.py ===
        initialize_schema(cursor)
        # ===========================================
        self.connection.commit()
        logger.info("Схема базы данных инициализирована успешно (вызов из base.py).")

    # - Вспомогательные методы для получения/создания ID -
    # (Примеры, реализация зависит от остального кода)

    def _get_or_create_project_id(self, project_name: str) -> Optional[int]:
        """Получает ID проекта или создает новый."""
        if not self.connection:
            logger.error("Нет активного соединения с БД.")
            return None

        cursor = self.connection.cursor()
        cursor.execute("SELECT id FROM projects WHERE name = ?", (project_name,))
        row = cursor.fetchone()
        if row:
            return row[0]
        else:
            try:
                created_at_iso = datetime.now().isoformat()
                cursor.execute(
                    "INSERT INTO projects (name, created_at) VALUES (?, ?)",
                    (project_name, created_at_iso)
                )
                self.connection.commit()
                project_id = cursor.lastrowid
                logger.info(f"Создан новый проект '{project_name}' (ID: {project_id}).")
                return project_id
            except sqlite3.Error as e:
                logger.error(f"Ошибка при создании проекта '{project_name}': {e}")
                self.connection.rollback()
                return None

    def _get_sheet_id_by_name(self, project_id: int, sheet_name: str) -> Optional[int]:
        """Получает ID листа по имени."""
        if not self.connection:
            logger.error("Нет активного соединения с БД.")
            return None

        cursor = self.connection.cursor()
        cursor.execute("SELECT id FROM sheets WHERE project_id = ? AND name = ?", (project_id, sheet_name))
        row = cursor.fetchone()
        return row[0] if row else None

    def _create_sheet_record(self, project_id: int, sheet_name: str, sheet_index: int, structure_json: str, raw_data_info_json: str) -> Optional[int]:
        """Создает запись о листе в БД."""
        if not self.connection:
            logger.error("Нет активного соединения с БД.")
            return None

        cursor = self.connection.cursor()
        try:
            cursor.execute(
                "INSERT INTO sheets (project_id, name, sheet_index, structure, raw_data_info) VALUES (?, ?, ?, ?, ?)",
                (project_id, sheet_name, sheet_index, structure_json, raw_data_info_json)
            )
            self.connection.commit()
            sheet_id = cursor.lastrowid
            logger.debug(f"Создан новый лист '{sheet_name}' (ID: {sheet_id}).")
            return sheet_id
        except sqlite3.Error as e:
            logger.error(f"Ошибка при создании записи листа '{sheet_name}': {e}")
            self.connection.rollback()
            return None

    # - Конец вспомогательных методов -

    # --- МЕТОДЫ ДЛЯ РАБОТЫ С СЫРЫМИ ДАННЫМИ ---

    def create_raw_data_table(self, sheet_name: str, column_names: List[str]) -> bool:
        """Создает таблицу в БД для хранения сырых данных листа."""
        if not self.connection:
            logger.error("Нет активного соединения для создания таблицы сырых данных.")
            return False
        return raw_data.create_raw_data_table(self.connection, sheet_name, column_names)

    def save_sheet_raw_data(self, sheet_id: int, sheet_name: str, raw_data_info: Dict[str, Any]) -> bool:
        """Сохраняет сырые данные листа."""
        if not self.connection:
            logger.error("Нет активного соединения для сохранения сырых данных.")
            return False
        return raw_data.save_sheet_raw_data(self.connection, sheet_id, sheet_name, raw_data_info)

    def load_sheet_raw_data(self, sheet_name: str) -> Dict[str, Any]:
        """Загружает сырые данные листа."""
        if not self.connection:
            logger.error("Нет активного соединения для загрузки сырых данных.")
            return {"column_names": [], "rows": []}
        return raw_data.load_sheet_raw_data(self.connection, sheet_name)

    # --- МЕТОДЫ ДЛЯ РАБОТЫ С РЕДАКТИРУЕМЫМИ ДАННЫМИ ---

    def load_sheet_editable_data(self, sheet_name: str) -> Dict[str, Any]:
        """Загружает редактируемые данные листа."""
        if not self.connection:
            logger.error("Нет активного соединения для загрузки редактируемых данных.")
            return {"column_names": [], "rows": []}
        return editable_data.load_sheet_editable_data(self.connection, sheet_name)

    def update_editable_cell(self, sheet_name: str, row_index: int, column_name: str, new_value: Any) -> bool:
        """Обновляет значение редактируемой ячейки."""
        if not self.connection:
            logger.error("Нет активного соединения для обновления редактируемой ячейки.")
            return False
        return editable_data.update_editable_cell(self.connection, sheet_name, row_index, column_name, new_value)

    # --- МЕТОДЫ ДЛЯ РАБОТЫ С ИСТОРИЕЙ ---

    def save_edit_history_record(
        self,
        project_id: int,
        sheet_id: Optional[int],
        cell_address: Optional[str],
        action_type: str,
        old_value: Any,
        new_value: Any,
        user: Optional[str] = None,
        details: Optional[dict] = None
    ) -> bool:
        """Сохраняет запись об изменении в истории."""
        if not self.connection:
            logger.error("Нет активного соединения для сохранения записи истории.")
            return False
        return history.save_edit_history_record(
            self.connection, project_id, sheet_id, cell_address,
            action_type, old_value, new_value, user, details
        )

    # --- МЕТОДЫ ДЛЯ РАБОТЫ С ФОРМУЛАМИ ---

    def save_sheet_formulas(self, sheet_id: int, formulas_data: List[Dict[str, Any]]) -> bool:
        """Сохраняет формулы для листа."""
        if not self.connection:
            logger.error("Нет соединения")
            return False
        return formulas.save_formulas(self.connection, sheet_id, formulas_data)

    def load_sheet_formulas(self, sheet_id: int) -> List[Dict[str, Any]]:
        """Загружает формулы для листа."""
        if not self.connection:
            logger.error("Нет соединения")
            return []
        return formulas.load_formulas(self.connection, sheet_id)

    # --- МЕТОДЫ ДЛЯ РАБОТЫ СО СТИЛЯМИ ---

    def save_sheet_styles(self, sheet_id: int, styled_ranges_data: List[Dict[str, Any]]) -> bool:
        """Сохраняет стили для листа."""
        if not self.connection:
            logger.error("Нет соединения")
            return False
        return styles.save_sheet_styles(self.connection, sheet_id, styled_ranges_data)

    def load_sheet_styles(self, sheet_id: int) -> List[Dict[str, Any]]:
        """Загружает стили для листа."""
        if not self.connection:
            logger.error("Нет соединения")
            return []
        return styles.load_sheet_styles(self.connection, sheet_id)

    # --- МЕТОДЫ ДЛЯ РАБОТЫ С ДИАГРАММАМИ ---

    def save_sheet_charts(self, sheet_id: int, charts_data: List[Dict[str, Any]]) -> bool:
        """Сохраняет диаграммы для листа."""
        if not self.connection:
            logger.error("Нет соединения")
            return False
        return charts.save_sheet_charts(self.connection, sheet_id, charts_data)

    def load_sheet_charts(self, sheet_id: int) -> List[Dict[str, Any]]:
        """Загружает диаграммы для листа."""
        if not self.connection:
            logger.error("Нет соединения")
            return []
        return charts.load_sheet_charts(self.connection, sheet_id)

    # --- МЕТОДЫ ДЛЯ РАБОТЫ С МЕТАДАННЫМИ ---

    def save_analysis_results(self, project_name: str, documentation_data: Dict[str, Any]):
        """Сохраняет результаты анализа документации в базу данных."""
        if not self.connection:
            logger.error("Нет активного соединения с БД для сохранения результатов.")
            return False  # Добавлен возврат значения

        # Этот метод пока оставим здесь, так как он координирует сохранение всех данных.
        # В будущем его можно будет разбить.

        try:
            cursor = self.connection.cursor()

            # 1. Создание/получение записи проекта
            project_id = self._get_or_create_project_id(project_name)
            if not project_id:
                return False

            # 2. Сохранение информации о листах и связанной информации
            sheets_data = documentation_data.get("sheets", {})
            for sheet_name, sheet_info in sheets_data.items():
                logger.debug(f"Обработка листа: {sheet_name}")

                # 2.1. Создание/получение записи листа
                sheet_index = sheet_info.get("index", 0)
                sheet_id = self._get_sheet_id_by_name(project_id, sheet_name)

                if not sheet_id:
                    structure_json = json.dumps(sheet_info.get("structure", []), cls=DateTimeEncoder, ensure_ascii=False)
                    raw_data_summary = {
                        "column_names": sheet_info.get("raw_data", {}).get("column_names", []),
                    }
                    raw_data_info_json = json.dumps(raw_data_summary, cls=DateTimeEncoder, ensure_ascii=False)
                    sheet_id = self._create_sheet_record(project_id, sheet_name, sheet_index, structure_json, raw_data_info_json)
                    if not sheet_id:
                        continue  # Пропускаем этот лист, если не удалось создать запись

                # 2.2. Сохранение формул листа
                formulas_data = sheet_info.get("formulas", [])
                self.save_sheet_formulas(sheet_id, formulas_data)

                # 2.3. Сохранение межлистовых ссылок (пока встроенная логика)
                cross_refs_data = sheet_info.get("cross_sheet_references", [])
                cursor.execute("DELETE FROM cross_sheet_references WHERE sheet_id = ?", (sheet_id,))
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
                logger.debug(f" Сохранено {len(cross_refs_data)} межлистовых ссылок для листа '{sheet_name}'.")

                # 2.4. Сохранение диаграмм
                charts_data = sheet_info.get("charts", [])
                self.save_sheet_charts(sheet_id, charts_data)

                # 2.5. Сохранение сырых данных
                raw_data_info_to_save = sheet_info.get("raw_data", {})
                if raw_data_info_to_save:
                    success = self.save_sheet_raw_data(sheet_id, sheet_name, raw_data_info_to_save)
                    if not success:
                        logger.error(f" Ошибка при сохранении сырых данных для листа '{sheet_name}'.")

                # === ИЗМЕНЕНИЕ: Создание и заполнение таблицы редактируемых данных ===
                # 2.6. Создание и заполнение таблицы редактируемых данных
                # Импортируем функцию внутри метода для избежания циклических импортов
                from src.storage.editable_data import create_and_populate_editable_table
                editable_data_created = create_and_populate_editable_table(
                    self.connection, sheet_id, sheet_name, raw_data_info_to_save
                )
                if not editable_data_created:
                    logger.error(f" Ошибка при создании/заполнении таблицы редактируемых данных для листа '{sheet_name}'.")
                else:
                     logger.debug(f" Таблица редактируемых данных для листа '{sheet_name}' создана/заполнена.")
                # ======================================================================

                # 2.7. Сохранение стилей
                styled_ranges_data = sheet_info.get("styled_ranges", [])
                if styled_ranges_data:
                    success = self.save_sheet_styles(sheet_id, styled_ranges_data)
                    if not success:
                        logger.error(f" Ошибка при сохранении стилей для листа '{sheet_name}'.")

                # 2.8. Сохранение объединенных ячеек (пока встроенная логика)
                merged_cells_data = sheet_info.get("merged_cells", [])
                if merged_cells_data:
                    cursor.execute("DELETE FROM merged_cells_ranges WHERE sheet_id = ?", (sheet_id,))
                    for range_addr in merged_cells_data:
                        if range_addr:
                            cursor.execute('''
                                INSERT OR IGNORE INTO merged_cells_ranges (sheet_id, range_address)
                                VALUES (?, ?)
                            ''', (sheet_id, range_addr))
                    self.connection.commit()
                    logger.debug(f" Сохранено {len(merged_cells_data)} объединенных диапазонов для листа '{sheet_name}'.")

            # Подтверждаем транзакцию
            self.connection.commit()
            logger.info("Все результаты анализа успешно сохранены в базу данных.")
            return True

        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при сохранении результатов анализа: {e}")
            self.connection.rollback()
            return False
        except Exception as e:
            logger.error(f"Неожиданная ошибка при сохранении результатов анализа: {e}", exc_info=True)
            self.connection.rollback()
            return False

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

            # Получаем имя проекта
            cursor.execute("SELECT name FROM projects LIMIT 1")
            project_row = cursor.fetchone()
            if not project_row:
                logger.warning("В базе данных не найдено информации о проекте.")
                return {}
            project_name = project_row[0]
            all_data["project_info"]["name"] = project_name
            logger.debug(f"Загруженное имя проекта: {project_name}")

            # Получаем список имен листов
            cursor.execute("SELECT name FROM sheets ORDER BY sheet_index")
            sheets_rows = cursor.fetchall()
            sheet_names = [row[0] for row in sheets_rows]
            logger.debug(f"Найденные листы: {sheet_names}")

            # Загружаем данные для каждого листа
            for sheet_name in sheet_names:
                logger.debug(f"Загрузка данных для листа: '{sheet_name}'")

                # Используем вспомогательный метод из metadata.py или оставляем локальный
                # sheet_data = metadata.load_sheet_data(self.connection, sheet_name) # Будущий вариант
                sheet_data = self._load_sheet_data(sheet_name)  # Текущий вариант

                if sheet_data:
                    all_data["sheets"][sheet_name] = sheet_data
                else:
                    logger.warning(f"Данные для листа '{sheet_name}' не были загружены или оказались пустыми.")

            logger.debug("Загрузка всех данных проекта завершена.")
            return all_data

        except Exception as e:
            logger.error(f"Ошибка при загрузке всех данных проекта: {e}", exc_info=True)
            return {}

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
            "raw_data": {},
            "formulas": [],
            "cross_sheet_references": [],
            "charts": [],
            "styled_ranges": [],
            "merged_cells": [],
        }

        if not self.connection:
            logger.error("Нет активного соединения с БД.")
            return None

        try:
            cursor = self.connection.cursor()

            # Получаем ID листа и другую информацию
            cursor.execute("SELECT id, structure, raw_data_info FROM sheets WHERE name = ?", (sheet_name,))
            sheet_row = cursor.fetchone()
            if not sheet_row:
                logger.warning(f"Информация о листе '{sheet_name}' не найдена в БД.")
                return None
            sheet_id, structure_json, raw_data_info_json = sheet_row
            logger.debug(f"ID листа '{sheet_name}': {sheet_id}")

            # Десериализуем структуру
            if structure_json:
                try:
                    sheet_data["structure"] = json.loads(structure_json)
                    logger.debug(f" Загружена структура листа '{sheet_name}' ({len(sheet_data['structure'])} элементов).")
                except json.JSONDecodeError as e:
                    logger.error(f"Ошибка десериализации структуры листа '{sheet_name}': {e}")

            # Загрузка информации о сырых данных
            if raw_data_info_json:
                try:
                    raw_data_summary = json.loads(raw_data_info_json)
                    sheet_data["raw_data"]["column_names"] = raw_data_summary.get("column_names", [])
                except json.JSONDecodeError as e:
                    logger.error(f"Ошибка десериализации сводки сырых данных листа '{sheet_name}': {e}")

            # Загружаем сами сырые данные из отдельной таблицы
            loaded_raw_data = self.load_sheet_raw_data(sheet_name)
            if loaded_raw_data and (loaded_raw_data.get("column_names") or loaded_raw_data.get("rows")):
                sheet_data["raw_data"] = loaded_raw_data
                logger.debug(f" Загружены полные сырые данные листа '{sheet_name}' ({len(loaded_raw_data.get('rows', []))} строк).")
            else:
                logger.debug(f" Полные сырые данные для листа '{sheet_name}' отсутствуют или пусты.")

            # Загружаем формулы
            sheet_data['formulas'] = self.load_sheet_formulas(sheet_id)
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

            # Загружаем диаграммы
            sheet_data['charts'] = self.load_sheet_charts(sheet_id)
            logger.debug(f" Загружено {len(sheet_data['charts'])} диаграмм для листа '{sheet_name}'.")

            # Загружаем стили
            sheet_data["styled_ranges"] = self.load_sheet_styles(sheet_id)
            logger.debug(f" Загружено {len(sheet_data['styled_ranges'])} стилей для листа '{sheet_name}'.")

            # Загружаем объединенные ячейки
            cursor.execute("SELECT range_address FROM merged_cells_ranges WHERE sheet_id = ?", (sheet_id,))
            merged_cells_rows = cursor.fetchall()
            sheet_data["merged_cells"] = [row[0] for row in merged_cells_rows]
            logger.debug(f" Загружено {len(sheet_data['merged_cells'])} объединенных диапазонов для листа '{sheet_name}'.")

            logger.debug(f"Данные листа '{sheet_name}' загружены успешно.")
            return sheet_data

        except Exception as e:
            logger.error(f"Ошибка при загрузке данных листа '{sheet_name}': {e}", exc_info=True)
            return None

    # - ТОЧКА ВХОДА ДЛЯ ТЕСТИРОВАНИЯ -

    if __name__ == "__main__":
        # Простой тест подключения и инициализации
        print("--- ТЕСТ ХРАНИЛИЩА (base.py) ---")

        # Определяем путь к тестовой БД относительно корня проекта
        # Используем project_root, определенный выше в файле
        test_db_path = project_root / "data" / "test_db_base.sqlite" 
        print(f"Путь к тестовой БД: {test_db_path}")

        try:
            # ИМПОРТ ВНУТРИ БЛОКА MAIN ДЛЯ ИЗБЕЖАНИЯ ПРЕДУПРЕЖДЕНИЙ PYLANCE
            # Явно импортируем класс из текущего модуля
            from src.storage.base import ProjectDBStorage
            
            # Теперь Pylance знает, откуда берется ProjectDBStorage
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
