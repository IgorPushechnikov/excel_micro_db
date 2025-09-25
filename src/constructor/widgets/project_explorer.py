# src/constructor/widgets/project_explorer.py
"""
Модуль для виджета обозревателя проекта.
"""

import logging
from typing import Optional, List, Dict, Any

from PySide6.QtCore import Qt, Slot, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem,
    QMessageBox
)

# Импортируем AppController
from src.core.app_controller import AppController

# Получаем логгер
logger = logging.getLogger(__name__)


class ProjectExplorer(QWidget):
    """
    Виджет для отображения структуры проекта.
    """

    # Сигнал, испускаемый при выборе листа
    sheet_selected = Signal(str) # Передаёт имя листа

    def __init__(self, app_controller: AppController):
        """
        Инициализирует обозреватель проекта.

        Args:
            app_controller (AppController): Экземпляр основного контроллера приложения.
        """
        super().__init__()
        self.app_controller: AppController = app_controller
        self._tree: Optional[QTreeWidget] = None
        self._setup_ui()
        self._load_project_structure()
        logger.debug("ProjectExplorer инициализирован.")

    def _setup_ui(self):
        """Настраивает UI виджета."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._tree = QTreeWidget()
        self._tree.setHeaderLabel("") # Или "Структура проекта"
        self._tree.setAlternatingRowColors(True)
        
        # Подключаем сигнал клика
        self._tree.itemClicked.connect(self._on_item_clicked)

        layout.addWidget(self._tree)
        logger.debug("UI ProjectExplorer настроено.")

    def _load_project_structure(self):
        """Загружает структуру проекта из AppController."""
        logger.debug("Загрузка структуры проекта в ProjectExplorer...")
        
        # Проверим, загружен ли проект
        if not self.app_controller.is_project_loaded:
             logger.info("Проект не загружен. ProjectExplorer будет пуст.")
             self._tree.clear()
             placeholder_item = QTreeWidgetItem(["<Нет открытого проекта>"])
             placeholder_item.setFlags(Qt.NoItemFlags) # Нельзя выбрать
             self._tree.addTopLevelItem(placeholder_item)
             return

        try:
            # Временно используем storage напрямую
            storage = self.app_controller.storage
            if storage and storage.connection:
                 sheets_data = storage.load_all_sheets_metadata()
                 logger.debug(f"Получены данные листов: {sheets_data}")

                 self._tree.clear()
                 if sheets_data:
                    for sheet_info in sheets_data:
                        sheet_name = sheet_info.get("name", "Безымянный лист")
                        sheet_id = sheet_info.get("sheet_id", "N/A")
                        
                        item = QTreeWidgetItem([sheet_name])
                        item.setData(0, Qt.UserRole, sheet_id) # Сохраняем ID
                        item.setData(0, Qt.UserRole + 1, sheet_name) # Сохраняем имя
                        
                        self._tree.addTopLevelItem(item)
                    
                    # Раскрываем все элементы
                    self._tree.expandAll()
                 else:
                     placeholder_item = QTreeWidgetItem(["<Нет листов в проекте>"])
                     placeholder_item.setFlags(Qt.NoItemFlags)
                     self._tree.addTopLevelItem(placeholder_item)
            else:
                 logger.warning("Нет доступа к storage AppController.")
                 placeholder_item = QTreeWidgetItem(["<Ошибка доступа к данным>"])
                 placeholder_item.setFlags(Qt.NoItemFlags)
                 self._tree.addTopLevelItem(placeholder_item)

        except Exception as e:
            logger.error(f"Ошибка при загрузке структуры проекта: {e}", exc_info=True)
            QMessageBox.warning(
                self, 
                "Ошибка загрузки проекта", 
                f"Не удалось загрузить структуру проекта:\n{e}"
            )

    @Slot(QTreeWidgetItem, int)
    def _on_item_clicked(self, item: QTreeWidgetItem, column: int):
        """Слот для обработки клика по элементу дерева."""
        # Проверяем, является ли элемент листом (по наличию данных)
        sheet_id = item.data(0, Qt.UserRole)
        sheet_name = item.data(0, Qt.UserRole + 1)
        
        if sheet_id is not None and sheet_name:
            logger.debug(f"Выбран лист: {sheet_name} (ID: {sheet_id})")
            # Испускаем сигнал
            self.sheet_selected.emit(sheet_name)
        else:
            logger.debug(f"Кликнут элемент, не являющийся листом: {item.text(0)}")


# --- Вспомогательные функции (если понадобятся) ---

# def _create_tree_item(sheet_info: Dict[str, Any]) -> QTreeWidgetItem:
#     """Создаёт элемент дерева для листа."""
#     item = QTreeWidgetItem([sheet_info['name']])
#     item.setData(0, Qt.UserRole, sheet_info['sheet_id'])
#     # Можно добавить иконки и т.д.
#     return item
