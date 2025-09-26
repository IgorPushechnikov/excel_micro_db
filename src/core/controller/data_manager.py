# src/core/controller/data_manager.py
"""
Модуль для управления данными листа.
Отвечает за загрузку, обновление и историю редактирования данных листа.
"""
import logging
from typing import Dict, Any, List, Optional, Tuple
import sqlite3  # Для аннотаций типов, если нужно

# Импортируем AppController из родительского пакета core для аннотаций типов и доступа к storage
from ..app_controller import AppController

# Импортируем ProjectDBStorage
from src.storage.base import ProjectDBStorage

# Импортируем logger из utils
from src.utils.logger import get_logger

logger = get_logger(__name__)

class DataManager:
    """
    Класс для управления данными листа.
    """

    def __init__(self, app_controller: AppController):
        """
        Инициализирует менеджер данных.

        Args:
            app_controller (AppController): Ссылка на основной контроллер приложения.
        """
        self.app_controller = app_controller
        logger.debug("DataManager инициализирован.")

    def get_sheet_data(self, sheet_name: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        Получает данные листа для отображения в GUI.

        Возвращает кортеж (raw_data, editable_data).

        Args:
            sheet_name (str): Имя листа.

        Returns:
            Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
            Кортеж из списков сырых и редактируемых данных.
        """
        storage = self.app_controller.storage
        if not storage:
            logger.error("Проект не загружен.")
            return ([], [])

        try:
            logger.debug(f"Загрузка данных для листа '{sheet_name}'...")

            # Загружаем "сырые данные"
            raw_data = storage.load_sheet_raw_data(sheet_name)
            logger.debug(f"Загружено {len(raw_data)} записей сырых данных.")

            # Получаем sheet_id для загрузки редактируемых данных
            sheet_id = self._get_sheet_id_by_name(sheet_name)
            if sheet_id is None:
                logger.warning(f"Не найден sheet_id для листа '{sheet_name}'. Редактируемые данные не загружены.")
                editable_data = []
            else:
                # Загружаем редактируемые данные
                # --- Используем исправленный метод из storage ---
                editable_data = storage.load_sheet_editable_data(sheet_id, sheet_name)
                logger.debug(f"Загружено {len(editable_data)} записей редактируемых данных.")

            return (raw_data, editable_data)
        except Exception as e:
            logger.error(f"Ошибка при загрузке данных для листа '{sheet_name}': {e}", exc_info=True)
            return ([], [])

    def get_sheet_editable_data(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """
        Получает редактируемые данные листа в формате, ожидаемом SheetDataModel.
        Этот метод вызывается из SheetEditor.

        Args:
            sheet_name (str): Имя листа.

        Returns:
            Optional[Dict[str, Any]]: Словарь с 'column_names' и 'rows', или None в случае ошибки.
        """
        storage = self.app_controller.storage
        if not storage:
            logger.error("Проект не загружен.")
            return None

        try:
            logger.debug(f"Загрузка редактируемых данных для листа '{sheet_name}' для GUI...")

            # Для MVP предположим, что имена столбцов - это стандартные имена Excel (A, B, C...)
            # В будущем их нужно будет загружать из метаданных листа
            # from src.constructor.components.sheet_editor import SheetDataModel # <-- УБРАНО
            # dummy_model = SheetDataModel({"column_names": [], "rows": []}) # <-- УБРАНО
            max_columns = 20  # Временное значение, в реальности нужно из метаданных
            column_names = self._generate_excel_column_names(max_columns) # <-- ИСПОЛЬЗУЕМ МЕТОД DataManager

            # Получаем sheet_id
            sheet_id = self._get_sheet_id_by_name(sheet_name)
            if sheet_id is None:
                logger.warning(f"Не найден sheet_id для листа '{sheet_name}'.")
                # Возвращаем структуру с пустыми данными
                return {"column_names": column_names, "rows": []}

            # Загружаем редактируемые данные
            editable_data_list = storage.load_sheet_editable_data(sheet_id, sheet_name)
            
            # Преобразуем список словарей в список кортежей для модели
            # Предполагаем, что данные хранятся по адресам, и нам нужно создать матрицу
            # Это упрощенная реализация для MVP
            rows = []
            if editable_data_list:
                # Найдем максимальный номер строки и столбца
                max_row = 0
                max_col = 0
                cell_dict = {}
                for item in editable_data_list:
                    addr = item["cell_address"]
                    # Простой парсер адреса (только для A1, B2 и т.д., без диапазонов)
                    col_part = ""
                    row_part = ""
                    for char in addr:
                        if char.isalpha():
                            col_part += char
                        else:
                            row_part += char
                    if row_part and col_part:
                        row_idx = int(row_part) - 1  # 0-based
                        col_idx = self._column_letter_to_index(col_part)
                        max_row = max(max_row, row_idx)
                        max_col = max(max_col, col_idx)
                        cell_dict[(row_idx, col_idx)] = item["value"]
                
                # Создаем матрицу данных
                for r in range(max_row + 1):
                    row = []
                    for c in range(max(max_col + 1, len(column_names))):
                        row.append(cell_dict.get((r, c), ""))
                    rows.append(tuple(row))
            else:
                # Если данных нет, создаем одну пустую строку
                rows = [tuple([""] * len(column_names))]

            result = {
                "column_names": column_names,
                "rows": rows
            }
            logger.debug(f"Редактируемые данные для листа '{sheet_name}' подготовлены для GUI.")
            return result

        except Exception as e:
            logger.error(f"Ошибка при подготовке редактируемых данных для листа '{sheet_name}': {e}", exc_info=True)
            return None

    def _column_letter_to_index(self, letter: str) -> int:
        """Преобразует букву столбца Excel (например, 'A', 'Z', 'AA') в 0-базовый индекс."""
        result = 0
        for char in letter:
            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
        return result - 1  # 0-based index

    def _get_sheet_id_by_name(self, sheet_name: str) -> Optional[int]:
        """Вспомогательный метод для получения sheet_id по имени листа."""
        storage = self.app_controller.storage
        if not storage or not storage.connection:
            return None

        try:
            cursor = storage.connection.cursor()
            # Предполагаем project_id = 1
            cursor.execute("SELECT sheet_id FROM sheets WHERE name = ? AND project_id = 1", (sheet_name,))
            result = cursor.fetchone()
            return result[0] if result else None
        except sqlite3.Error:
            return None

    def update_sheet_cell_in_project(self, sheet_name: str, row_index: int, column_name: str, new_value: str) -> bool:
        """
        Обновляет значение ячейки в проекте. Этот метод вызывается из SheetEditor.
        Он преобразует координаты из GUI в адрес ячейки Excel и вызывает update_cell_value.

        Args:
            sheet_name (str): Имя листа.
            row_index (int): Индекс строки (0-based).
            column_name (str): Имя столбца (например, 'A', 'B', 'AA').
            new_value (str): Новое значение ячейки.

        Returns:
            bool: True, если обновление успешно, иначе False.
        """
        try:
            # Преобразуем координаты в адрес ячейки Excel (например, A1)
            cell_address = f"{column_name}{row_index + 1}"  # row_index 0-based -> Excel 1-based
            return self.update_cell_value(sheet_name, cell_address, new_value)
        except Exception as e:
            logger.error(f"Ошибка при обновлении ячейки в проекте: {e}", exc_info=True)
            return False

    def update_cell_value(self, sheet_name: str, cell_address: str, new_value: Any) -> bool:
        """
        Обновляет значение ячейки и записывает изменение в историю.

        Args:
            sheet_name (str): Имя листа.
            cell_address (str): Адрес ячейки (например, 'A1').
            new_value (Any): Новое значение ячейки.

        Returns:
            bool: True, если обновление успешно, иначе False.
        """
        storage = self.app_controller.storage
        if not storage:
            logger.error("Проект не загружен. Невозможно обновить ячейку.")
            return False

        try:
            logger.debug(f"Обновление ячейки {cell_address} на листе '{sheet_name}' на значение '{new_value}'.")

            # 1. Получаем sheet_id
            sheet_id = self._get_sheet_id_by_name(sheet_name)
            if sheet_id is None:
                logger.error(f"Не найден sheet_id для листа '{sheet_name}'. Обновление невозможно.")
                return False

            # 2. Получаем старое значение (для истории)
            # Это может быть сложно, так как оно может быть в raw_data или editable_data
            # Для простоты MVP, предположим, что мы можем получить его из editable_data
            # или установить None. Более точная логика может потребоваться.
            old_value = None  # TODO: Получить реальное старое значение

            # 3. Обновляем редактируемые данные
            # --- Используем исправленный метод из storage ---
            if not storage.update_editable_cell(sheet_id, sheet_name, cell_address, new_value):
                logger.error(f"Не удалось обновить редактируемую ячейку {cell_address} на листе '{sheet_name}'.")
                return False

            # 4. Записываем в историю редактирования
            # --- Используем метод из storage ---
            if not storage.save_edit_history_record(sheet_id, cell_address, old_value, new_value):
                logger.warning(f"Не удалось записать изменение ячейки {cell_address} в историю.")

            logger.info(f"Ячейка {cell_address} на листе '{sheet_name}' успешно обновлена.")
            return True
        except Exception as e:
            logger.error(f"Ошибка при обновлении ячейки {cell_address} на листе '{sheet_name}': {e}", exc_info=True)
            return False

    def get_edit_history(self, sheet_name: Optional[str] = None, limit: Optional[int] = 10) -> List[Dict[str, Any]]:
        """
        Получает историю редактирования.

        Args:
            sheet_name (Optional[str]): Имя листа для фильтрации. Если None, вся история.
            limit (Optional[int]): Максимальное количество записей.

        Returns:
            List[Dict[str, Any]]: Список записей истории.
        """
        storage = self.app_controller.storage
        if not storage:
            logger.error("Проект не загружен.")
            return []

    def _generate_excel_column_names(self, num_cols: int) -> List[str]:
        """
        Генерирует список имён столбцов Excel (A, B, ..., Z, AA, AB, ...).

        Args:
            num_cols (int): Количество столбцов.

        Returns:
            List[str]: Список имён столбцов.
        """
        if num_cols <= 0:
            return []
        names = []
        for i in range(num_cols):
            name = ""
            j = i
            while j >= 0:
                name = chr(j % 26 + ord('A')) + name
                j = j // 26 - 1
            names.append(name)
        return names

    # ... (остальной код get_edit_history без изменений)
        try:
            sheet_id = self._get_sheet_id_by_name(sheet_name) if sheet_name else None
            return storage.load_edit_history(sheet_id, limit)
        except Exception as e:
            logger.error(f"Ошибка при загрузке истории редактирования: {e}", exc_info=True)
            return []

