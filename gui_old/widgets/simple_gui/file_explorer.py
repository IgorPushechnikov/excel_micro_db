# backend/constructor/widgets/simple_gui/file_explorer.py
"""
Проводник Excel файла для упрощённого GUI - показывает имя файла и список листов.
"""
from PySide6.QtWidgets import QDockWidget, QListWidget, QListWidgetItem
from PySide6.QtCore import Qt, Slot, Signal
from pathlib import Path


class ExcelFileExplorer(QDockWidget):
    """Проводник Excel файла - показывает имя файла и список листов."""
    
    sheet_selected = Signal(str)
    
    def __init__(self, parent=None):
        super().__init__("Excel файл", parent)
        self.file_name = ""
        self.sheets_list = []
        
        # Создаем виджет для отображения
        self.list_widget = QListWidget()
        self.setWidget(self.list_widget)
        
        # Подключаем сигнал выбора
        self.list_widget.itemClicked.connect(self._on_item_clicked)
        
        # Устанавливаем разрешенные области
        self.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)
    
    def load_excel_file(self, file_path: str, sheet_names: list):
        """Загружает информацию о Excel файле."""
        self.file_name = Path(file_path).name
        self.sheets_list = sheet_names
        
        # Обновляем заголовок док-виджета
        self.setWindowTitle(f"Excel файл: {self.file_name}")
        
        # Очищаем и заполняем список
        self.list_widget.clear()
        for sheet_name in sheet_names:
            item = QListWidgetItem(sheet_name)
            self.list_widget.addItem(item)
    
    def clear_file(self):
        """Очищает информацию о файле."""
        self.file_name = ""
        self.sheets_list = []
        self.list_widget.clear()
        self.setWindowTitle("Excel файл")
    
    @Slot()
    def _on_item_clicked(self, item):
        """Обработчик клика по элементу списка."""
        sheet_name = item.text()
        self.sheet_selected.emit(sheet_name)