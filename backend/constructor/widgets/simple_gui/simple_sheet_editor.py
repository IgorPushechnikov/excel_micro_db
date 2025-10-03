# backend/constructor/widgets/simple_gui/simple_sheet_editor.py
"""
Упрощённый табличный редактор для отображения данных листа Excel.
"""
from typing import Optional
from PySide6.QtWidgets import QWidget, QVBoxLayout, QTableView, QLabel
from PySide6.QtCore import Qt
import logging

from backend.utils.logger import get_logger
from backend.constructor.widgets.simple_gui.simple_sheet_model import SimpleSheetModel

logger = get_logger(__name__)


class SimpleSheetEditor(QWidget):
    """Виджет для отображения содержимого листа Excel."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db_path = None
        self.sheet_name = None
        self._model: Optional[SimpleSheetModel] = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        
        self.label_sheet_name = QLabel("Лист: <Не выбран>")
        self.label_sheet_name.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(self.label_sheet_name)

        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectItems)
        self.table_view.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        
        # Настройка заголовков
        horizontal_header = self.table_view.horizontalHeader()
        horizontal_header.setSectionResizeMode(horizontal_header.ResizeMode.Interactive)
        vertical_header = self.table_view.verticalHeader()
        vertical_header.setSectionResizeMode(vertical_header.ResizeMode.Fixed)
        vertical_header.setDefaultSectionSize(20)

        layout.addWidget(self.table_view)
    
    def load_sheet(self, db_path: str, sheet_name: str):
        """Загружает данные листа из БД и отображает их."""
        logger.info(f"Загрузка листа '{sheet_name}' из БД: {db_path}")
        self.db_path = db_path
        self.sheet_name = sheet_name
        self.label_sheet_name.setText(f"Лист: {sheet_name}")
        
        # Создаём новую модель
        self._model = SimpleSheetModel(db_path, sheet_name)
        
        # Устанавливаем модель в представление
        self.table_view.setModel(self._model)
        
        logger.info(f"Лист '{sheet_name}' успешно загружен в редактор")
    
    def clear_sheet(self):
        """Очищает отображение."""
        logger.debug("Очистка редактора листа")
        self.db_path = None
        self.sheet_name = None
        self.label_sheet_name.setText("Лист: <Не выбран>")
        self.table_view.setModel(None)
        self._model = None