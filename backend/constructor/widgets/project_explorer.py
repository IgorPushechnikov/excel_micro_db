# backend/constructor/widgets/project_explorer.py
"""
Виджет "Обозреватель проекта" для Excel Micro DB GUI.
Отображает структуру открытого проекта в виде дерева.
"""

# === ИЗМЕНЕНО: Импорт QWidget вместо QDockWidget ===
from typing import Optional, Dict, Any, List
from pathlib import Path

# === ИЗМЕНЕНО: Импорт необходимых виджетов ===
from PySide6.QtWidgets import (
    QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QLabel # Убран QDockWidget, добавлен QLabel
)
# Импортируем QtCore.Qt для доступа к перечислениям и Signal/Slot
from PySide6.QtCore import Qt, Signal, Slot
# Явно импортируем sqlite3 в заголовке файла, чтобы Pylance был доволен
import sqlite3 # <-- ДОБАВЛЕНО

# ИСПРАВЛЕНО: Импорт logger теперь из utils внутри backend
from utils.logger import get_logger

logger = get_logger(__name__)

# Типы элементов дерева для внутреннего использования
ITEM_TYPE_PROJECT = 1001
ITEM_TYPE_SHEETS_FOLDER = 1002
ITEM_TYPE_SHEET = 1003
# ITEM_TYPE_FORMULAS_FOLDER = 1004 # Удалено
# ITEM_TYPE_STYLES_FOLDER = 1005   # Удалено
# ITEM_TYPE_CHARTS_FOLDER = 1006   # Удалено


# === ИЗМЕНЕНО: Класс ProjectExplorer наследуется от QWidget ===
class ProjectExplorer(QWidget):
    """
    Виджет для отображения структуры проекта (только листов).
    Ранее наследовался от QDockWidget, теперь является обычным QWidget,
    помещаемым в QDockWidget внешним кодом (например, MainWindow).
    """

    # Сигнал, испускаемый при выборе элемента листа
    sheet_selected = Signal(str) # Передаёт имя листа

    def __init__(self, parent=None):
        # === ИЗМЕНЕНО: Вызов конструктора QWidget ===
        # super().__init__("Обозреватель проекта", parent) # <-- СТАРОЕ: Вызов QDockWidget.__init__
        super().__init__(parent) # <-- НОВОЕ: Вызов QWidget.__init__
        self.project_data: Optional[Dict[str, Any]] = None
        self.db_path: Optional[str] = None
        self._setup_ui()

    def _setup_ui(self):
        """Настройка пользовательского интерфейса."""
        # === ИЗМЕНЕНО: Layout напрямую для self (QWidget) ===
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0) # Убираем отступы

        # === ИЗМЕНЕНО: QLabel для заголовка внутри виджета (опционально) ===
        # self.label_sheet_name = QLabel("Листы") # <-- ИЗМЕНЕНО С "Структура проекта"
        # self.label_sheet_name.setStyleSheet("font-weight: bold; padding: 5px;")
        # layout.addWidget(self.label_sheet_name)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("") # <-- УДАЛЕНО ЗАГОЛОВОК "Структура проекта"
        self.tree.setAlternatingRowColors(True)

        # Отключаем стандартный контекстное меню, если оно не нужно пока
        # self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        # self.tree.customContextMenuRequested.connect(self._on_context_menu)

        # === ИСПРАВЛЕНО: Подключение сигнала выбора элемента ===
        # Подключаем сигнал изменения текущего выбранного элемента
        self.tree.currentItemChanged.connect(self._on_item_changed)

        layout.addWidget(self.tree)

    def load_project(self, project_data: Dict[str, Any], db_path: str):
        """
        Загрузка и отображение структуры проекта (только листов).

        Args:
            project_data (Dict[str, Any]): Данные проекта из AppController.
            db_path (str): Путь к файлу БД проекта.
        """
        logger.debug("Загрузка проекта в ProjectExplorer")
        self.project_data = project_data
        self.db_path = db_path
        self.tree.clear()

        if not self.project_data:
            logger.warning("Попытка загрузить пустые данные проекта")
            return

        try:
            project_name = self.project_data.get("project_name", "Проект")

            # --- Создание корневого элемента проекта ---
            project_item = QTreeWidgetItem(self.tree, [f"{project_name}"])
            project_item.setData(0, Qt.ItemDataRole.UserRole, ITEM_TYPE_PROJECT) # <-- ИСПРАВЛЕНО
            # project_item.setIcon(0, QIcon("path/to/project_icon.png")) # TODO: Иконка проекта
            project_item.setExpanded(True) # Раскрываем проект по умолчанию

            # --- Создание папки "Листы" ---
            sheets_folder_item = QTreeWidgetItem(project_item, ["Листы"])
            sheets_folder_item.setData(0, Qt.ItemDataRole.UserRole, ITEM_TYPE_SHEETS_FOLDER) # <-- ИСПРАВЛЕНО
            # sheets_folder_item.setIcon(0, QIcon("path/to/sheets_icon.png")) # TODO: Иконка папки
            sheets_folder_item.setExpanded(True) # Раскрываем папку листов по умолчанию

            # --- Заполнение листов ---
            # Получаем список листов из БД
            sheets_info = self._get_sheets_info_from_db()
            if sheets_info:
                for sheet_info in sheets_info:
                    sheet_name = sheet_info.get("name", "Безымянный лист")
                    sheet_item = QTreeWidgetItem(sheets_folder_item, [sheet_name])
                    sheet_item.setData(0, Qt.ItemDataRole.UserRole, ITEM_TYPE_SHEET) # <-- ИСПРАВЛЕНО
                    # === ИСПРАВЛЕНО: Сохраняем имя листа в данных элемента (UserRole + 1) ===
                    sheet_item.setData(0, Qt.ItemDataRole.UserRole + 1, sheet_name) # Сохраняем имя листа в данных элемента
                    # sheet_item.setIcon(0, QIcon("path/to/sheet_icon.png")) # TODO: Иконка листа
            else:
                # Если листов нет или ошибка, показываем заглушку
                no_sheets_item = QTreeWidgetItem(sheets_folder_item, ["(Нет данных)"])
                # === ИСПРАВЛЕНО: Установка флагов ===
                no_sheets_item.setFlags(Qt.ItemFlag.NoItemFlags) # <-- ИСПРАВЛЕНО # Делаем недоступным для выбора

            # --- Удалены папки "Формулы", "Стили", "Диаграммы" ---
            # --- Создание других папок (заглушки) ---
            # Они будут заполняться позже

            logger.info(f"Структура проекта '{project_name}' загружена в обозреватель")

        except Exception as e:
            logger.error(f"Ошибка при загрузке структуры проекта в обозреватель: {e}", exc_info=True)
            # Показываем сообщение об ошибке в дереве или статус баре родителя
            error_item = QTreeWidgetItem(self.tree, [f"Ошибка загрузки: {e}"])
            # === ИСПРАВЛЕНО: Установка флагов ===
            error_item.setFlags(Qt.ItemFlag.NoItemFlags) # <-- ИСПРАВЛЕНО

    def _get_sheets_info_from_db(self) -> List[Dict[str, Any]]:
        """
        Получает информацию о листах из БД проекта.
        Это временная реализация, в будущем можно использовать AppController.
        """
        if not self.db_path:
            logger.error("Путь к БД проекта не установлен")
            return []

        # Явно импортируем sqlite3 внутри метода, чтобы Pylance был уверен в его наличии
        # import sqlite3 # <-- УДАЛЕНО, так как уже импортирован в заголовке

        try:
            logger.debug(f"Подключение к БД проекта: {self.db_path}")
            # === ИСПРАВЛЕНО: Использование контекстного менеджера для соединения ===
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row  # Для доступа по именам колонок
                cursor = conn.cursor()

                # Запрашиваем информацию о листах
                cursor.execute("SELECT id, name, sheet_index FROM sheets ORDER BY sheet_index")
                rows = cursor.fetchall()

                sheets = [dict(row) for row in rows]
                logger.debug(f"Из БД получено {len(sheets)} листов")
                return sheets

        # === ИСПРАВЛЕНО: Уточнение типа исключения ===
        except sqlite3.Error as e:  # <-- Теперь Pylance знает, что это sqlite3.Error
            logger.error(f"Ошибка работы с БД проекта при получении листов: {e}")
            # QMessageBox.warning(self, "Ошибка БД", f"Не удалось получить список листов из БД проекта:\n{e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка при получении листов из БД: {e}", exc_info=True)
            return []

    def clear_project(self):
        """Очистка отображения проекта."""
        logger.debug("Очистка обозревателя проекта")
        self.project_data = None
        self.db_path = None
        self.tree.clear()

    # === ДОБАВЛЕНО: Метод _on_item_changed ===
    @Slot(object, object) # <-- Используем object для типов аргументов, так как они сложные
    def _on_item_changed(self, current: QTreeWidgetItem, previous: QTreeWidgetItem):
        """Обработчик изменения текущего выбранного элемента."""
        # Игнорируем previous, он нам не нужен
        _ = previous

        if current is None:
            return

        # === ИСПРАВЛЕНО: Получение типа элемента ===
        item_type = current.data(0, Qt.ItemDataRole.UserRole)  # <-- ИСПРАВЛЕНО

        # Проверяем, является ли выбранный элемент листом
        if item_type == ITEM_TYPE_SHEET:
            # === ИСПРАВЛЕНО: Получение имени листа из правильных данных ===
            sheet_name = current.data(0, Qt.ItemDataRole.UserRole + 1)  # <-- ИСПРАВЛЕНО (UserRole + 1)
            if sheet_name:
                logger.debug(f"Выбран лист: {sheet_name}")
                # Испускаем сигнал с именем выбранного листа
                self.sheet_selected.emit(sheet_name)
            else:
                logger.warning("Выбран элемент листа, но имя листа не найдено в данных элемента")
