# src/constructor/widgets/sheet_editor.py
"""
Модуль для виджета редактора листа с Fluent Widgets.
"""

import logging
from typing import Optional, List, Dict, Any

# --- НОВОЕ: Импорт из qfluentwidgets ---
from qfluentwidgets import TableWidget
# --- КОНЕЦ НОВОГО ---

# Импортируем AppController
from src.core.app_controller import AppController

# Получаем логгер
logger = logging.getLogger(__name__)


# --- НОВОЕ: SheetEditor наследуется от TableWidget ---
class SheetEditor(TableWidget):
    """
    Виджет для отображения и редактирования содержимого листа с Fluent дизайном.
    Наследуется от TableWidget из qfluentwidgets.
    """

    def __init__(self, app_controller: AppController):
        """
        Инициализирует редактор листа.

        Args:
            app_controller (AppController): Экземпляр основного контроллера приложения.
        """
        super().__init__()
        self.app_controller: AppController = app_controller
        self.sheet_name: Optional[str] = None
        # self._table_view: Optional[QTableView] = None # Уже есть self (TableWidget)
        # TODO: Добавить модель данных (QAbstractItemModel)
        # self._model = None
        self._setup_ui()
        logger.debug("SheetEditor (Fluent) инициализирован.")

    def _setup_ui(self):
        """Настраивает UI виджета."""
        logger.debug("Настройка UI SheetEditor (Fluent)...")
        # TableWidget уже настроен, просто установим параметры
        self.setAlternatingRowColors(True)
        # Настройка заголовков
        horizontal_header = self.horizontalHeader()
        # horizontal_header.setSectionResizeMode(QHeaderView.Interactive) # По умолчанию Interactive
        vertical_header = self.verticalHeader()
        # vertical_header.setSectionResizeMode(QHeaderView.Fixed) # По умолчанию Fixed
        vertical_header.setDefaultSectionSize(20) # Высота строки
        logger.debug("UI SheetEditor (Fluent) настроено.")

    def load_sheet(self, sheet_name: str):
        """
        Загружает данные листа для отображения.

        Args:
            sheet_name (str): Имя листа для загрузки.
        """
        logger.info(f"Загрузка листа '{sheet_name}' в SheetEditor (Fluent)...")
        self.sheet_name = sheet_name
        # self._label_sheet_name.setText(f"Лист: {sheet_name}") # Убираем, так как это TableWidget

        # TODO: Получить данные листа из AppController
        # Например: self.app_controller.get_sheet_editable_data(sheet_name)
        # Или: self.app_controller.storage.load_sheet_raw_data(sheet_name) + формулы + стили
        # Пока используем заглушку
        
        # --- ЗАГЛУШКА ---
        # sheet_data = {
        #     "column_names": ["A", "B", "C"],
        #     "rows": [
        #         ("Заголовок1", "Заголовок2", "Заголовок3"),
        #         ("Данные1", 10, 3.14),
        #         ("Данные2", 20, 2.71),
        #     ]
        # }
        # --- КОНЕЦ ЗАГЛУШКИ ---
        
        if not self.app_controller or not self.app_controller.is_project_loaded:
            logger.warning("AppController не готов или проект не загружен.")
            QMessageBox.warning(self, "Ошибка", "Проект не загружен.")
            return

        try:
            # Попробуем получить данные через AppController
            # Предполагаем, что у AppController будет метод, возвращающий данные в нужном формате
            # editable_data = self.app_controller.get_sheet_editable_data(sheet_name)
            
            # Пока используем storage напрямую
            storage = self.app_controller.storage
            if storage and storage.connection:
                 # Заглушка: загружаем только сырые данные
                 raw_data = storage.load_sheet_raw_data(sheet_name)
                 logger.debug(f"Загружены сырые данные для листа '{sheet_name}': {len(raw_data) if raw_data else 0} записей.")
                 
                 # TODO: Создать модель данных и привязать её к self (TableWidget)
                 # self._model = SomeDataModel(raw_data) # Пока нет модели
                 # self.setModel(self._model)
                 
                 # Временное решение: просто покажем сообщение
                 QMessageBox.information(self, "Заглушка", f"Лист '{sheet_name}' выбран.\nДанные: {len(raw_data) if raw_data else 0} записей.\n(Редактор пока в разработке)")
            else:
                 logger.error("Нет доступа к storage AppController.")
                 QMessageBox.critical(self, "Ошибка", "Нет доступа к данным проекта.")
        except Exception as e:
            logger.error(f"Ошибка при загрузке листа '{sheet_name}': {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить лист '{sheet_name}':\n{e}")


# --- Вспомогательные классы (например, модель данных) будут добавлены позже ---
# class SheetDataModel(QAbstractTableModel):
#     ...
