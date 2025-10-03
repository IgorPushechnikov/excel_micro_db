# backend/constructor/widgets/simple_gui/simple_sheet_editor.py
"""
Упрощённый табличный редактор для отображения данных листа Excel.
"""
from typing import Optional
from pathlib import Path
import logging
from PySide6.QtWidgets import QWidget, QVBoxLayout, QTableView, QLabel
from PySide6.QtCore import Qt

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

        # Настройка логгера для SimpleSheetEditor
        self.editor_logger = self._setup_logger()

        self._setup_ui()

    def _setup_logger(self) -> logging.Logger:
        """Настраивает логгер для записи в файл проекта."""
        logger_name = f"SimpleSheetEditor"
        logger_instance = logging.getLogger(logger_name)
        logger_instance.setLevel(logging.DEBUG)

        # Очищаем существующие хендлеры, чтобы избежать дублирования
        logger_instance.handlers.clear()

        # Определяем путь к папке проекта и папке логов
        # Используем db_path, если он известен, иначе логируем в основной лог
        if self.db_path:
            db_path_obj = Path(self.db_path)
            project_dir = db_path_obj.parent
            logs_dir = project_dir / "logs"
            logs_dir.mkdir(exist_ok=True)  # Создаем папку logs, если она не существует
            log_file_path = logs_dir / "simple_sheet_editor.log"
        else:
            # Если db_path неизвестен, используем общий лог
            log_file_path = Path("logs") / "simple_sheet_editor.log"
            Path("logs").mkdir(exist_ok=True)

        # Создаем FileHandler
        file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)

        # Форматтер для логов
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)

        logger_instance.addHandler(file_handler)
        return logger_instance

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
        self.editor_logger.info(f"Загрузка листа '{sheet_name}' из БД: {db_path}")
        logger.info(f"Загрузка листа '{sheet_name}' из БД: {db_path}") # Лог в основном логе тоже оставим
        self.db_path = db_path
        self.sheet_name = sheet_name
        self.label_sheet_name.setText(f"Лист: {sheet_name}")

        # Обновляем логгер, чтобы он писал в правильную папку проекта
        # (теперь, когда db_path известен)
        self.editor_logger = self._setup_logger()

        # Создаём новую модель
        self.editor_logger.debug(f"Создание новой SimpleSheetModel для '{sheet_name}'")
        try:
            self._model = SimpleSheetModel(db_path, sheet_name)
            self.editor_logger.debug(f"SimpleSheetModel создана, rowCount: {self._model.rowCount()}, columnCount: {self._model.columnCount()}")
        except Exception as e:
            self.editor_logger.error(f"Ошибка при создании SimpleSheetModel: {e}", exc_info=True)
            logger.error(f"Ошибка при создании SimpleSheetModel: {e}", exc_info=True)
            return # Не устанавливаем модель в случае ошибки

        # Устанавливаем модель в представление
        self.table_view.setModel(self._model)
        self.editor_logger.info(f"Лист '{sheet_name}' успешно загружен в редактор")

    def clear_sheet(self):
        """Очищает отображение."""
        logger.debug("Очистка редактора листа")
        self.editor_logger.debug("Очистка редактора листа")
        self.db_path = None
        self.sheet_name = None
        self.label_sheet_name.setText("Лист: <Не выбран>")
        self.table_view.setModel(None)
        self._model = None