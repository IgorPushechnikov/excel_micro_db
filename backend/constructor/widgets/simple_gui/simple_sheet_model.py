# backend/constructor/widgets/simple_gui/simple_sheet_model.py
"""
Модель данных для упрощённого табличного редактора.
Загружает данные и стили из БД и предоставляет их QTableView.
"""
import sqlite3
import json
import re
from typing import Dict, Any, List, Optional
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor, QFont
import logging

from backend.utils.logger import get_logger

logger = get_logger(__name__)


class SimpleSheetModel(QAbstractTableModel):
    """Модель данных для отображения содержимого листа Excel."""
    
    def __init__(self, db_path: str, sheet_name: str, parent=None):
        """
        Инициализирует модель с данными из БД.
        
        Args:
            db_path (str): Путь к файлу БД.
            sheet_name (str): Имя листа для загрузки.
            parent: Родительский объект.
        """
        super().__init__(parent)
        self.db_path = db_path
        self.sheet_name = sheet_name
        
        # Данные таблицы
        self._data: List[List[Any]] = []
        # Стили ячеек: {(row, col): {"font_color": "#FF0000", "bg_color": "#FFFFFF", ...}}
        self._cell_styles: Dict[tuple, Dict[str, Any]] = {}
        
        # Заголовки столбцов Excel (A, B, C...)
        self._column_headers = []
        
        self._load_data_from_db()
        self._generate_column_headers()
    
    def _load_data_from_db(self):
        """Загружает данные и стили из БД."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 1. Получаем ID листа
            cursor.execute("SELECT id FROM sheets WHERE name = ?", (self.sheet_name,))
            sheet_row = cursor.fetchone()
            if not sheet_row:
                logger.warning(f"Лист '{self.sheet_name}' не найден в БД")
                return
            sheet_id = sheet_row[0]
            
            # 2. Загружаем "сырые" данные
            cursor.execute(
                "SELECT cell_address, value FROM raw_data WHERE sheet_name = ? ORDER BY cell_address",
                (self.sheet_name,)
            )
            raw_data_rows = cursor.fetchall()
            
            # 3. Загружаем стили
            cursor.execute(
                "SELECT range_address, style_attributes FROM sheet_styles WHERE sheet_id = ?",
                (sheet_id,)
            )
            style_rows = cursor.fetchall()
            
            conn.close()
            
            # Обработка данных
            max_row = 0
            max_col = 0
            data_map = {}
            
            for addr, val in raw_data_rows:
                row_idx, col_idx = self._parse_cell_address(addr)
                if row_idx is not None and col_idx is not None:
                    data_map[(row_idx, col_idx)] = val
                    max_row = max(max_row, row_idx)
                    max_col = max(max_col, col_idx)
            
            # Создание 2D списка
            self._data = [[""] * (max_col + 1) for _ in range(max_row + 1)]
            for (r, c), val in data_map.items():
                if 0 <= r <= max_row and 0 <= c <= max_col:
                    self._data[r][c] = val
            
            # Обработка стилей
            for range_addr, style_attrs_json in style_rows:
                style_attrs = json.loads(style_attrs_json)
                self._apply_style_to_range(range_addr, style_attrs)
            
            logger.info(f"Загружено {len(self._data)} строк данных и {len(self._cell_styles)} стилей для листа '{self.sheet_name}'")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки данных из БД для листа '{self.sheet_name}': {e}", exc_info=True)
    
    def _parse_cell_address(self, addr: str) -> tuple:
        """Парсит адрес ячейки (например, A1) в индексы (row, col)."""
        try:
            col_part = ''.join(filter(str.isalpha, addr)).upper()
            row_part = ''.join(filter(str.isdigit, addr))
            
            if not row_part or not col_part:
                return None, None
            
            row_idx = int(row_part) - 1  # 1-based to 0-based
            col_idx = self._column_letter_to_index(col_part)
            
            return row_idx, col_idx
        except:
            return None, None
    
    def _column_letter_to_index(self, letter: str) -> int:
        """Преобразует букву столбца Excel (например, 'A', 'Z', 'AA') в 0-базовый индекс."""
        result = 0
        for char in letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1  # 0-based index
    
    def _apply_style_to_range(self, range_addr: str, style_attrs: Dict[str, Any]):
        """Применяет стиль к диапазону ячеек."""
        # Парсер адреса диапазона Excel (например, A1:B2)
        range_pattern = re.compile(r'^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$')
        match = range_pattern.match(range_addr)
        if not match:
            logger.warning(f"Не удалось распознать формат адреса '{range_addr}'. Пропущено.")
            return
        
        start_col_letter, start_row_str, end_col_letter, end_row_str = match.groups()
        if end_col_letter is None or end_row_str is None:
            end_col_letter = start_col_letter
            end_row_str = start_row_str
        
        try:
            start_col_index = self._column_letter_to_index(start_col_letter)
            start_row_index = int(start_row_str) - 1
            end_col_index = self._column_letter_to_index(end_col_letter)
            end_row_index = int(end_row_str) - 1

            for r in range(start_row_index, end_row_index + 1):
                for c in range(start_col_index, end_col_index + 1):
                    self._cell_styles[(r, c)] = style_attrs
        except (ValueError, TypeError) as e:
            logger.warning(f"Ошибка преобразования индексов для '{range_addr}': {e}")
    
    def _generate_column_headers(self):
        """Генерирует имена столбцов Excel (A, B, ..., Z, AA, AB...)."""
        if not self._data:
            self._column_headers = []
            return
        
        num_cols = len(self._data[0]) if self._data else 0
        self._column_headers = []
        for i in range(num_cols):
            name = ""
            temp = i
            while temp >= 0:
                name = chr(temp % 26 + ord('A')) + name
                temp = temp // 26 - 1
                if temp < 0:
                    break
            self._column_headers.append(name if name else "A")
    
    # Реализация QAbstractTableModel
    def rowCount(self, parent=QModelIndex()):
        return len(self._data)
    
    def columnCount(self, parent=QModelIndex()):
        return len(self._data[0]) if self._data else 0
    
    def data(self, index: QModelIndex, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        
        row = index.row()
        col = index.column()
        
        if row >= len(self._data) or col >= len(self._data[0]):
            return None
        
        if role == Qt.ItemDataRole.DisplayRole:
            return self._data[row][col]
        elif role == Qt.ItemDataRole.BackgroundRole:
            style = self._cell_styles.get((row, col))
            if style:
                bg_color = style.get("fill_fg_color") or style.get("fill_bg_color")
                if bg_color and isinstance(bg_color, str) and bg_color.startswith('#'):
                    return QBrush(QColor(bg_color))
        elif role == Qt.ItemDataRole.ForegroundRole:
            style = self._cell_styles.get((row, col))
            if style:
                font_color = style.get("font_color")
                if font_color and isinstance(font_color, str) and font_color.startswith('#'):
                    return QBrush(QColor(font_color))
        elif role == Qt.ItemDataRole.FontRole:
            style = self._cell_styles.get((row, col))
            if style:
                font = QFont()
                 # Жирность
                font_b = style.get("font_b")
                if font_b in (True, '1', 'true', 'True'):
                    font.setBold(True)
                # Курсив
                font_i = style.get("font_i")
                if font_i in (True, '1', 'true', 'True'):
                    font.setItalic(True)
                # Подчеркивание
                font_u = style.get("font_u")
                if font_u and font_u != "none":
                    font.setUnderline(True)
                # Зачеркивание
                font_strike = style.get("font_strike")
                if font_strike in (True, '1', 'true', 'True'):
                    font.setStrikeOut(True)
                return font
        
        return None
    
    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                if 0 <= section < len(self._column_headers):
                    return self._column_headers[section]
            elif orientation == Qt.Orientation.Vertical:
                return str(section + 1)  # Номера строк 1-based
        return None