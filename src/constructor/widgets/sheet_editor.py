# src/constructor/widgets/sheet_editor.py
"""
Виджет-редактор для отображения и редактирования содержимого листа Excel.
"""
import sys
from typing import Optional, Dict, Any, List, NamedTuple, Union
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTableView, QLabel, QMessageBox,
    QAbstractItemView, QHeaderView, QApplication, QMenu
)
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, Slot, Signal, QPersistentModelIndex
from PySide6.QtGui import QBrush, QColor, QAction
import sqlite3
import logging

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from src.core.app_controller import AppController

from src.utils.logger import get_logger

logger = get_logger(__name__)

# === НОВОЕ: Структура для хранения информации об одном редактировании ===
class EditAction(NamedTuple):
    """Представляет одно действие редактирования для Undo/Redo."""
    row: int
    col: int
    old_value: Any
    new_value: Any
# =====================================================================

# === ИЗМЕНЕНО: SheetDataModel с новым сигналом ===
class SheetDataModel(QAbstractTableModel):
    """
    Модель данных для отображения и редактирования содержимого листа в QTableView.
    """
    # === НОВОЕ: Сигнал, испускаемый ДО изменения данных ===
    # QModelIndex, старое значение, новое значение
    cellDataAboutToChange = Signal(QModelIndex, object, object)
    # ===================================================
    
    # === СУЩЕСТВУЮЩИЙ: Сигнал для внутреннего использования ===
    dataChangedExternally = Signal(QModelIndex, QModelIndex)
    # =========================================================

    def __init__(self, editable_data: Dict[str, Any], parent=None):
        super().__init__(parent)
        self._editable_data = editable_data
        self._column_names = self._editable_data.get("column_names", [])
        self._rows = self._editable_data.get("rows", [])

    def rowCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = QModelIndex()) -> int:
        return len(self._rows)

    def columnCount(self, parent: Union[QModelIndex, QPersistentModelIndex] = QModelIndex()) -> int:
        return len(self._column_names)

    # ИСПРАВЛЕНО: Тип параметра index для совместимости с базовым классом
    def data(self, index: Union[QModelIndex, QPersistentModelIndex], role=Qt.ItemDataRole.DisplayRole):
        # Преобразуем QPersistentModelIndex в QModelIndex если нужно
        if isinstance(index, QPersistentModelIndex):
            index = QModelIndex(index)
            
        if not index.isValid():
            return None
        row = index.row()
        col = index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            if row < len(self._rows) and col < len(self._rows[row]):
                value = self._rows[row][col]
                return str(value) if value is not None else ""
        elif role == Qt.ItemDataRole.ToolTipRole:
            if row < len(self._rows) and col < len(self._rows[row]):
                value = self._rows[row][col]
                return f"Значение: {repr(value)}"
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal and section < len(self._column_names):
                return self._column_names[section]
            elif orientation == Qt.Orientation.Vertical:
                return str(section + 1)
        return None

    # ИСПРАВЛЕНО: Тип параметра index для совместимости с базовым классом
    def flags(self, index: Union[QModelIndex, QPersistentModelIndex]) -> Qt.ItemFlag:
        # Преобразуем QPersistentModelIndex в QModelIndex если нужно
        if isinstance(index, QPersistentModelIndex):
            index = QModelIndex(index)
            
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled

    # === ИЗМЕНЕНО: setData с испусканием сигнала до изменения ===
    # ИСПРАВЛЕНО: Тип параметра index для совместимости с базовым классом
    def setData(self, index: Union[QModelIndex, QPersistentModelIndex], value, role=Qt.ItemDataRole.EditRole):
        """
        Устанавливает данные в модель. Вызывается, когда пользователь редактирует ячейку.
        Испускает cellDataAboutToChange до изменения и dataChanged после.
        """
        # Преобразуем QPersistentModelIndex в QModelIndex если нужно
        if isinstance(index, QPersistentModelIndex):
            index = QModelIndex(index)
            
        if index.isValid() and role == Qt.ItemDataRole.EditRole:
            row = index.row()
            col = index.column()

            if row < len(self._rows) and col < len(self._rows[row]):
                # 1. Получаем старое значение
                old_value = self._rows[row][col]
                # 2. Преобразуем новое значение к строке, как это делается в data()
                new_value_str = str(value) if value is not None else ""

                # 3. Испускаем сигнал ДО изменения
                # Это позволяет подписчикам (например, SheetEditor) узнать
                # детали изменения заранее и, например, записать его в историю Undo.
                logger.debug(f"SheetDataModel: Испускание cellDataAboutToChange для [{row},{col}]: '{old_value}' -> '{new_value_str}'")
                self.cellDataAboutToChange.emit(index, old_value, new_value_str)

                # 4. Обновляем данные в модели
                self._rows[row][col] = new_value_str

                # 5. Сообщаем представлению, что данные изменились
                # Это обновит отображение ячейки
                self.dataChanged.emit(index, index, [role])
                
                logger.debug(f"SheetDataModel: Данные ячейки [{row}, {col}] изменены с '{old_value}' на '{new_value_str}'.")
                # Возвращаем True, чтобы показать, что изменение принято
                return True
        # Если индекс невалиден или роль не подходит, возвращаем False
        return False
    # =========================================================

    def setDataInternal(self, row: int, col: int, value):
        """Внутренне обновляет данные модели без вызова setData."""
        if 0 <= row < len(self._rows) and 0 <= col < len(self._rows[row]):
            self._rows[row][col] = value
            index = self.index(row, col)
            if index.isValid():
                self.dataChangedExternally.emit(index, index)
                logger.debug(f"Модель (внутр.): Данные ячейки [{row}, {col}] обновлены до '{value}'.")

# === SheetEditor с подключением к новому сигналу ===
class SheetEditor(QWidget):
    """
    Виджет для редактирования/просмотра содержимого одного листа.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.project_db_path: Optional[str] = None
        self.sheet_name: Optional[str] = None
        self._model: Optional[SheetDataModel] = None
        self.app_controller: Optional['AppController'] = None
        
        # === Стеки для Undo/Redo ===
        self._undo_stack: List[EditAction] = []
        self._redo_stack: List[EditAction] = []
        self._max_undo_steps = 50
        # ==========================
        
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.label_sheet_name = QLabel("Лист: <Не выбран>")
        self.label_sheet_name.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(self.label_sheet_name)

        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.ContiguousSelection)
        
        # Настройка контекстного меню
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self._on_context_menu)

        horizontal_header = self.table_view.horizontalHeader()
        horizontal_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        
        vertical_header = self.table_view.verticalHeader()
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        vertical_header.setDefaultSectionSize(20)

        layout.addWidget(self.table_view)

        # Создание действий для Undo/Redo
        self.action_undo = QAction("Отменить", self)
        self.action_undo.setShortcut("Ctrl+Z")
        self.action_undo.triggered.connect(self.undo)
        self.action_undo.setEnabled(False)

        self.action_redo = QAction("Повторить", self)
        self.action_redo.setShortcut("Ctrl+Y")
        self.action_redo.triggered.connect(self.redo)
        self.action_redo.setEnabled(False)

    @Slot(object)
    def _on_context_menu(self, position):
        """Создает и показывает контекстное меню для таблицы."""
        context_menu = QMenu(self)
        context_menu.addAction(self.action_undo)
        context_menu.addAction(self.action_redo)
        context_menu.exec(self.table_view.viewport().mapToGlobal(position))

    def set_app_controller(self, controller: Optional['AppController']):
        self.app_controller = controller
        logger.debug("SheetEditor: AppController установлен/обновлён.")

    def load_sheet(self, project_db_path: str, sheet_name: str):
        logger.info(f"Загрузка листа '{sheet_name}' из БД: {project_db_path}")
        self.project_db_path = project_db_path
        self.sheet_name = sheet_name
        self.label_sheet_name.setText(f"Лист: {sheet_name}")

        # Очистка стеков Undo/Redo при загрузке нового листа
        self._clear_undo_redo_stacks()

        if not self.app_controller:
            logger.error("SheetEditor: AppController не установлен для загрузки листа.")
            QMessageBox.critical(self, "Ошибка", "Контроллер приложения не доступен для загрузки данных.")
            return

        try:
            editable_data = self.app_controller.get_sheet_editable_data(sheet_name)
            
            if editable_data is not None and 'column_names' in editable_data:
                self._model = SheetDataModel(editable_data)
                # === НОВОЕ: Подключение к новому сигналу модели ===
                # Подключаем сигнал модели о предстоящем изменении к слоту редактора
                self._model.cellDataAboutToChange.connect(self._on_cell_data_about_to_change)
                # =================================================
                self._model.dataChanged.connect(self._on_model_data_changed)
                self._model.dataChangedExternally.connect(self.table_view.dataChanged)
                self.table_view.setModel(self._model)
                logger.info(f"Лист '{sheet_name}' успешно загружен в редактор. "
                            f"Строк: {len(editable_data.get('rows', []))}, "
                            f"Столбцов: {len(editable_data.get('column_names', []))}")
            else:
                self.table_view.setModel(None)
                self._model = None
                logger.warning(f"Редактируемые данные для листа '{sheet_name}' не найдены или пусты.")
        except Exception as e:
            logger.error(f"Ошибка при загрузке листа '{sheet_name}': {e}", exc_info=True)
            QMessageBox.critical(
                self, 
                "Ошибка загрузки", 
                f"Не удалось загрузить содержимое листа '{sheet_name}':\n{e}"
            )
            self.table_view.setModel(None)
            self._model = None

    # === НОВОЕ: Слот для обработки сигнала о предстоящем изменении ===
    @Slot(QModelIndex, object, object)
    def _on_cell_data_about_to_change(self, index: QModelIndex, old_value, new_value):
        """
        Слот, вызываемый, когда модель сигнализирует о предстоящем изменении данных.
        Здесь формируется EditAction и добавляется в стек Undo.
        """
        if not index.isValid():
            return
        row = index.row()
        col = index.column()
        logger.debug(f"SheetEditor: Получено уведомление об изменении [{row},{col}]: '{old_value}' -> '{new_value}'")

        # Создаем действие для Undo
        action = EditAction(row=row, col=col, old_value=old_value, new_value=new_value)
        # Добавляем в стек Undo
        self._push_to_undo_stack(action)
    # =============================================================

    @Slot(QModelIndex, QModelIndex)
    def _on_model_data_changed(self, top_left: QModelIndex, bottom_right: QModelIndex):
        """
        Слот, вызываемый, когда модель сигнализирует об изменении данных.
        Здесь происходит вызов AppController для сохранения изменений.
        """
        logger.debug(f"SheetEditor._on_model_data_changed вызван для диапазона: ({top_left.row()},{top_left.column()}) - ({bottom_right.row()},{bottom_right.column()})")

        if not self.app_controller or not self.sheet_name or not self._model:
            logger.error("SheetEditor: Не хватает компонентов для обработки изменений.")
            return

        if not hasattr(self.app_controller, 'update_sheet_cell_in_project'):
             logger.error("SheetEditor: AppController не имеет метода 'update_sheet_cell_in_project'.")
             return

        # Обрабатываем только одиночные ячейки
        if top_left == bottom_right and top_left.isValid():
            row = top_left.row()
            col = top_left.column()
            
            if not (0 <= row < self._model.rowCount() and 0 <= col < self._model.columnCount()):
                logger.warning(f"SheetEditor: Изменение за пределами модели. row={row}, col={col}")
                return

            column_name = self._model._column_names[col] if col < len(self._model._column_names) else f"Col_{col}"
            # Новое значение уже в модели, получаем его
            new_value = self._model._rows[row][col] if row < len(self._model._rows) and col < len(self._model._rows[row]) else None
            
            logger.debug(f"SheetEditor: Обнаружено изменение в ячейке [{row}, {column_name}]. Новое значение: '{new_value}'")
            
            try:
                # ИСПРАВЛЕНО: Приведение new_value к str для соответствия сигнатуре метода
                success = self.app_controller.update_sheet_cell_in_project(
                    sheet_name=self.sheet_name,
                    row_index=row,
                    column_name=column_name,
                    new_value=str(new_value) if new_value is not None else "" # Приведение к str
                )
                
                if success:
                    logger.info(f"Изменение в ячейке [{self.sheet_name}][{row}, {column_name}] успешно сохранено в БД и истории.")
                    # --- УДАЛЕНО: добавление в стек здесь ---
                    # self._push_to_undo_stack(...) - теперь делается в _on_cell_data_about_to_change
                    # ----------------------------------------
                else:
                    logger.error(f"Не удалось сохранить изменение в ячейке [{self.sheet_name}][{row}, {column_name}] в БД.")
                    QMessageBox.warning(self, "Ошибка сохранения", 
                                        f"Не удалось сохранить изменение в ячейке {column_name}{row+1}.\n"
                                        "Изменение в интерфейсе будет сохранено до закрытия, но не записано в проект.")
                    
            except Exception as e:
                critical_error_msg = f"Исключение при сохранении изменения в ячейке [{self.sheet_name}][{row}, {column_name}]: {e}"
                logger.error(critical_error_msg, exc_info=True)
                QMessageBox.critical(self, "Критическая ошибка сохранения", 
                                    f"Произошла ошибка при попытке сохранить изменение:\n{e}\n\n"
                                    "Изменение в интерфейсе будет сохранено до закрытия, но не записано в проект.")
        else:
            range_info = f"от ({top_left.row()},{top_left.column()}) до ({bottom_right.row()},{bottom_right.column()})"
            logger.debug(f"SheetEditor: Обнаружен диапазон изменений {range_info}. Обработка диапазонов не реализована.")

    # === Методы Undo/Redo ===
    def undo(self):
        """Отменяет последнее действие редактирования."""
        if not self._undo_stack or not self._model:
            logger.debug("SheetEditor: Нечего отменять или модель не доступна.")
            return

        action = self._undo_stack.pop()
        logger.debug(f"SheetEditor: Отмена действия: [{action.row}, {action.col}] '{action.new_value}' -> '{action.old_value}'")

        # Восстанавливаем старое значение в модели (используя внутренний метод)
        self._model.setDataInternal(action.row, action.col, action.old_value)
        
        # Помещаем действие в стек Redo
        self._redo_stack.append(action)
        if len(self._redo_stack) > self._max_undo_steps:
            self._redo_stack.pop(0)
        
        self._update_undo_redo_actions()
        logger.info(f"Отменено изменение в ячейке [{action.row}, {action.col}].")

    def redo(self):
        """Повторяет последнее отмененное действие."""
        if not self._redo_stack or not self._model:
            logger.debug("SheetEditor: Нечего повторять или модель не доступна.")
            return

        action = self._redo_stack.pop()
        logger.debug(f"SheetEditor: Повтор действия: [{action.row}, {action.col}] '{action.old_value}' -> '{action.new_value}'")

        # Восстанавливаем новое значение в модели
        self._model.setDataInternal(action.row, action.col, action.new_value)
        
        # Помещаем действие обратно в стек Undo
        self._undo_stack.append(action)
        if len(self._undo_stack) > self._max_undo_steps:
            self._undo_stack.pop(0)
        
        self._update_undo_redo_actions()
        logger.info(f"Повторено изменение в ячейке [{action.row}, {action.col}].")

    def _push_to_undo_stack(self, action: EditAction):
        """Добавляет действие в стек отмены и очищает стек повторов."""
        self._undo_stack.append(action)
        if len(self._undo_stack) > self._max_undo_steps:
            self._undo_stack.pop(0)
        # Очищаем стек повторов при новом действии
        self._redo_stack.clear()
        self._update_undo_redo_actions()
        logger.debug(f"Действие добавлено в стек Undo: {action}")

    def _clear_undo_redo_stacks(self):
        """Очищает стеки Undo и Redo."""
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._update_undo_redo_actions()
        logger.debug("Стеки Undo/Redo очищены.")

    def _update_undo_redo_actions(self):
        """Обновляет состояние действий Undo/Redo."""
        self.action_undo.setEnabled(len(self._undo_stack) > 0)
        self.action_redo.setEnabled(len(self._redo_stack) > 0)
    # ===========================

    def clear_sheet(self):
        logger.debug("Очистка редактора листа")
        self.project_db_path = None
        self.sheet_name = None
        self._clear_undo_redo_stacks()
        self.label_sheet_name.setText("Лист: <Не выбран>")
        self.table_view.setModel(None)
        self._model = None
