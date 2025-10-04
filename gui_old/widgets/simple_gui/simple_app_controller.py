# backend/constructor/widgets/simple_gui/simple_app_controller.py
"""
Упрощённый контроллер приложения для упрощённого GUI.
Работает напрямую с БД, без концепции проекта.
"""
import sqlite3
from typing import Optional, Dict, Any, List, Tuple
from pathlib import Path
import logging

from backend.utils.logger import get_logger

logger = get_logger(__name__)


class SimpleAppController:
    """
    Упрощённый контроллер, который предоставляет только необходимые методы для SheetEditor.
    Работает с одной БД напрямую.
    """
    
    def __init__(self, db_path: str):
        """
        Инициализирует контроллер с указанием пути к БД.
        
        Args:
            db_path (str): Путь к файлу БД SQLite.
        """
        self.db_path = db_path
        self.connection: Optional[sqlite3.Connection] = None
        logger.debug(f"SimpleAppController инициализирован с БД: {db_path}")
    
    def initialize(self) -> bool:
        """
        Инициализирует подключение к БД.
        
        Returns:
            bool: True, если подключение успешно.
        """
        try:
            self.connection = sqlite3.connect(self.db_path)
            logger.info(f"Подключение к БД установлено: {self.db_path}")
            return True
        except sqlite3.Error as e:
            logger.error(f"Ошибка подключения к БД {self.db_path}: {e}")
            return False
    
    def get_sheet_editable_data(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """
        Получает редактируемые данные листа из БД.
        
        Args:
            sheet_name (str): Имя листа Excel.
        
        Returns:
            Optional[Dict[str, Any]]: Словарь с ключами 'rows' и 'column_names'.
        """
        if not self.connection:
            logger.error("Нет подключения к БД для получения данных листа")
            return None
        
        try:
            # Получаем ID листа
            cursor = self.connection.cursor()
            cursor.execute("SELECT id FROM sheets WHERE name = ?", (sheet_name,))
            result = cursor.fetchone()
            if not result:
                logger.warning(f"Лист '{sheet_name}' не найден в БД")
                return None
            
            sheet_id = result[0]
            
            # Загружаем "сырые" данные для этого листа
            # Используем raw_data таблицу, как в storage.raw_data
            # Структура: cell_address, value, value_type
            cursor.execute(
                "SELECT cell_address, value FROM raw_data WHERE sheet_name = ? ORDER BY cell_address",
                (sheet_name,)
            )
            raw_data_rows = cursor.fetchall()
            
            if not raw_data_rows:
                logger.info(f"Для листа '{sheet_name}' не найдено данных")
                return {"rows": [], "column_names": []}
            
            # Преобразуем raw_data в формат, подходящий для SheetDataModel
            # SheetDataModel ожидает список кортежей (или списков) для rows
            # и отдельно column_names (которые генерируются в модели)
            
            # Для простоты, будем создавать 2D список, где ячейки заполняются значениями по адресу
            # Сначала определим размеры таблицы
            max_row = 0
            max_col = 0
            data_map = {}
            
            for addr, val in raw_data_rows:
                # Парсим адрес ячейки (например, A1 -> row=0, col=0)
                col_part = ''.join(filter(str.isalpha, addr)).upper()
                row_part = ''.join(filter(str.isdigit, addr))
                
                if not row_part or not col_part:
                    continue
                
                row_idx = int(row_part) - 1  # 1-based to 0-based
                col_idx = self._column_letter_to_index(col_part)
                
                max_row = max(max_row, row_idx)
                max_col = max(max_col, col_idx)
                
                data_map[(row_idx, col_idx)] = val
            
            # Создаём 2D список
            rows = []
            for r in range(max_row + 1):
                row_data = []
                for c in range(max_col + 1):
                    row_data.append(data_map.get((r, c), ""))
                rows.append(row_data)
            
            logger.info(f"Получено {len(rows)} строк данных для листа '{sheet_name}'")
            return {"rows": rows}
            
        except sqlite3.Error as e:
            logger.error(f"Ошибка SQLite при загрузке данных листа '{sheet_name}': {e}")
            return None
        except Exception as e:
            logger.error(f"Неожиданная ошибка при загрузке данных листа '{sheet_name}': {e}", exc_info=True)
            return None
    
    def _column_letter_to_index(self, letter: str) -> int:
        """Преобразует букву столбца Excel (например, 'A', 'Z', 'AA') в 0-базовый индекс."""
        result = 0
        for char in letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1  # 0-based index
    
    def shutdown(self):
        """Закрывает соединение с БД."""
        if self.connection:
            self.connection.close()
            logger.info("Соединение с БД закрыто")
        else:
            logger.debug("Соединение с БД отсутствовало при shutdown")