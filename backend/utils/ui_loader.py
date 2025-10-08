# backend/utils/ui_loader.py
import logging
import yaml
from pathlib import Path
from typing import Dict, Any, Optional
from PySide6.QtWidgets import QMenu
from PySide6.QtGui import QAction # <-- QAction перемещён сюда
from PySide6.QtCore import QObject, Signal
from PySide6.QtGui import QKeySequence

from backend.utils.logger import get_logger

logger = get_logger(__name__)

class UICommandHandler(QObject):
    """
    Класс для хранения и вызова обработчиков команд из YAML.
    Позволяет связать строковое имя обработчика с реальным методом объекта.
    """
    def __init__(self, parent_obj: QObject):
        super().__init__()
        self._handlers: Dict[str, callable] = {}
        self._parent_obj = parent_obj

    def register_handler(self, handler_name: str, handler_func: callable):
        """Регистрирует обработчик команды."""
        self._handlers[handler_name] = handler_func
        logger.debug(f"Зарегистрирован обработчик команды: {handler_name}")

    def execute_handler(self, handler_name: str, *args, **kwargs):
        """Выполняет зарегистрированный обработчик."""
        handler_func = self._handlers.get(handler_name)
        if handler_func:
            try:
                logger.debug(f"Выполняется обработчик команды: {handler_name}")
                return handler_func(*args, **kwargs)
            except Exception as e:
                logger.error(f"Ошибка при выполнении обработчика '{handler_name}': {e}", exc_info=True)
        else:
            logger.error(f"Обработчик команды '{handler_name}' не найден.")


def load_context_menu_from_yaml(yaml_path: str, parent_widget: QObject, command_handler: UICommandHandler) -> Optional[QMenu]:
    """
    Загружает описание контекстного меню из YAML-файла и создаёт QMenu.

    Args:
        yaml_path (str): Путь к YAML-файлу.
        parent_widget (QObject): Виджет-родитель для QAction.
        command_handler (UICommandHandler): Объект для выполнения обработчиков команд.

    Returns:
        QMenu: Созданное контекстное меню или None в случае ошибки.
    """
    try:
        yaml_file_path = Path(yaml_path)
        if not yaml_file_path.exists():
            logger.error(f"Файл описания меню не найден: {yaml_file_path}")
            return None

        with open(yaml_file_path, 'r', encoding='utf-8') as f:
            menu_config = yaml.safe_load(f)

        if not menu_config or 'context_menu' not in menu_config:
            logger.error(f"Некорректный формат YAML-файла меню: {yaml_file_path}")
            return None

        menu = QMenu(parent_widget)

        for item_config in menu_config['context_menu']:
            action_id = item_config.get('id')
            action_text = item_config.get('text')
            action_shortcut = item_config.get('shortcut')
            handler_name = item_config.get('handler')

            if not action_id or not action_text or not handler_name:
                logger.warning(f"Пропущен пункт меню из-за отсутствия id, text или handler: {item_config}")
                continue

            action = QAction(action_text, parent_widget)
            if action_shortcut:
                action.setShortcut(QKeySequence(action_shortcut))

            # Связываем сигнал triggered с вызовом обработчика через command_handler
            # Используем lambda для захвата handler_name
            action.triggered.connect(lambda checked, h=handler_name: command_handler.execute_handler(h))

            menu.addAction(action)

        logger.info(f"Контекстное меню загружено из {yaml_file_path}")
        return menu

    except yaml.YAMLError as e:
        logger.error(f"Ошибка разбора YAML-файла меню {yaml_file_path}: {e}")
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке меню из {yaml_file_path}: {e}", exc_info=True)

    return None

# --- Пример использования (временно, для тестирования) ---
# if __name__ == "__main__":
#     import sys
#     from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton
#
#     app = QApplication(sys.argv)
#
#     # Пример виджета, который будет "обрабатывать" команды
#     class DummyWidget(QWidget):
#         def __init__(self):
#             super().__init__()
#             self.command_handler = UICommandHandler(self)
#             self.command_handler.register_handler("dummy_handler", self.on_dummy_action)
#
#             layout = QVBoxLayout(self)
#             btn = QPushButton("Показать меню (тест)")
#             btn.setContextMenuPolicy(Qt.CustomContextMenu)
#             btn.customContextMenuRequested.connect(self.show_menu)
#             layout.addWidget(btn)
#
#         def show_menu(self, pos):
#             menu = load_context_menu_from_yaml("test_menu.yaml", self, self.command_handler)
#             if menu:
#                 menu.exec_(self.sender().mapToGlobal(pos))
#
#         def on_dummy_action(self):
#             print("Выполнено действие dummy_handler!")
#
#     w = DummyWidget()
#     w.show()
#     sys.exit(app.exec())
