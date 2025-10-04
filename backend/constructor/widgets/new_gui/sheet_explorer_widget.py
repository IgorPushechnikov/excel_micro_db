# backend/constructor/widgets/new_gui/sheet_explorer_widget.py
"""
Виджет для отображения и управления списком листов (Sheet Explorer).
"""

import logging
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QListWidget, QListWidgetItem, QMenu, QMessageBox
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QCursor

from backend.utils.logger import get_logger

logger = get_logger(__name__)


class SheetExplorerWidget(QWidget):
    """
    Виджет обозревателя листов.
    
    Сигналы:
        sheet_selected(str): Сигнал, испускаемый при выборе листа.
        sheet_renamed(str, str): Сигнал, испускаемый при успешном переименовании листа.
                               (старое_имя, новое_имя)
    """
    sheet_selected = Signal(str)
    sheet_renamed = Signal(str, str)

    def __init__(self, app_controller, parent=None):
        """
        Инициализирует виджет обозревателя листов.

        Args:
            app_controller: Экземпляр AppController для взаимодействия с логикой приложения.
            parent: Родительский виджет.
        """
        super().__init__(parent)
        self.app_controller = app_controller
        self._setup_ui()
        self._setup_connections()

    def _setup_ui(self):
        """Создает элементы интерфейса."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0) # Убираем лишние отступы

        self.list_widget = QListWidget(self)
        self.list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        # Подключаем сигнал itemChanged здесь, он будет работать для пользовательских изменений
        self.list_widget.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.list_widget)

        self.setLayout(layout)

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        self.list_widget.itemSelectionChanged.connect(self._on_item_selection_changed)
        self.list_widget.itemDoubleClicked.connect(self._on_start_rename)
        self.list_widget.customContextMenuRequested.connect(self._on_custom_context_menu)

    def update_sheet_list(self):
        """
        Обновляет список листов из AppController.
        Важно: временно отключает сигнал itemChanged, чтобы избежать ложных срабатываний
        при программном заполнении списка.
        """
        self.list_widget.itemChanged.disconnect(self._on_item_changed)
        self.list_widget.clear()
        if not self.app_controller.is_project_loaded:
            self.list_widget.itemChanged.connect(self._on_item_changed)
            return

        try:
            sheet_names = self.app_controller.get_sheet_names()
            logger.debug(f"Обновление списка листов в обозревателе: {sheet_names}")
            for name in sheet_names:
                item = QListWidgetItem(name)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable) # Разрешаем редактирование
                # Сохраняем оригинальное имя для будущих проверок
                item.setData(Qt.ItemDataRole.UserRole, name)
                self.list_widget.addItem(item)
        except Exception as e:
            logger.error(f"Ошибка при обновлении списка листов в обозревателе: {e}", exc_info=True)
            # Можно добавить сообщение об ошибке в список
        
        # Включаем сигнал обратно
        self.list_widget.itemChanged.connect(self._on_item_changed)

    def _on_item_selection_changed(self):
        """Обработчик изменения выделения в списке."""
        selected_items = self.list_widget.selectedItems()
        if selected_items:
            sheet_name = selected_items[0].text()
            self.sheet_selected.emit(sheet_name)
        else:
            # Если ничего не выбрано, можно эмитить пустую строку или игнорировать
            pass

    def _on_start_rename(self, item: QListWidgetItem):
        """
        Обработчик начала переименования (например, по двойному клику).
        """
        if item:
            self.list_widget.editItem(item)

    def _on_custom_context_menu(self, position):
        """
        Обработчик контекстного меню.
        """
        item = self.list_widget.itemAt(position)
        if item is not None:
            menu = QMenu(self)
            rename_action = menu.addAction("Переименовать")
            action = menu.exec(QCursor.pos())
            if action == rename_action:
                self.list_widget.editItem(item)

    def _on_item_changed(self, item: QListWidgetItem):
        """
        Обработчик завершения редактирования элемента.
        Вызывается, когда пользователь завершает редактирование (Enter, потеря фокуса).
        """
        if not item:
            return

        old_name = item.data(Qt.ItemDataRole.UserRole)
        new_name = item.text()

        # Проверка на пустое имя или отсутствие изменений
        if not new_name.strip():
            QMessageBox.warning(self, "Ошибка переименования", "Имя листа не может быть пустым.")
            item.setText(old_name) # Восстанавливаем старое имя
            return

        if new_name == old_name:
            return # Ничего не изменилось

        # Пытаемся переименовать через AppController
        try:
            success = self.app_controller.rename_sheet(old_name, new_name)
            if success:
                logger.info(f"Лист успешно переименован: '{old_name}' -> '{new_name}'")
                # Обновляем сохраненное старое имя
                item.setData(Qt.ItemDataRole.UserRole, new_name)
                self.sheet_renamed.emit(old_name, new_name)
            else:
                QMessageBox.critical(self, "Ошибка переименования", f"Не удалось переименовать лист '{old_name}'.")
                item.setText(old_name) # Восстанавливаем старое имя
        except Exception as e:
            logger.error(f"Исключение при переименовании листа '{old_name}' в '{new_name}': {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка переименования", f"Ошибка при переименовании: {e}")
            item.setText(old_name) # Восстанавливаем старое имя
