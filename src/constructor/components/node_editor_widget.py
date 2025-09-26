# src/constructor/components/node_editor_widget.py
"""
Модуль для виджета нодового редактора (NodeGraphQt) нового GUI Excel Micro DB.
Интегрирует NodeGraphQt в виджет PySide6 для использования в QDockWidget.
"""

import logging
from typing import Optional, Dict, Any, List

from PySide6.QtWidgets import QWidget, QVBoxLayout, QMessageBox
from PySide6.QtCore import Qt, Signal

# --- Импорт NodeGraphQt ---
# ВАЖНО: Убедитесь, что NodeGraphQt совместим с PySide6.
# Возможно, потребуется установить NodeGraphQt-PySide6 или аналогичный пакет.
try:
    # Попытка импорта стандартного NodeGraphQt
    # NodeGraphQt может использовать PySide2 внутри, что может вызвать конфликты.
    # Если возникают ошибки импорта, проверьте версию NodeGraphQt или используйте совместимую версию.
    import NodeGraphQt
    from NodeGraphQt import NodeGraph, BaseNode
    from NodeGraphQt.constants import PortTypeEnum
    logger = logging.getLogger(__name__)
except ImportError as e:
    logger = logging.getLogger(__name__) # Получаем логгер, если NodeGraphQt не импортировался
    logger.error(f"Не удалось импортировать NodeGraphQt: {e}")
    NodeGraphQt = None
    NodeGraph = object # Заглушка для аннотаций типов
    BaseNode = object # Заглушка для аннотаций типов


# --- Пользовательские узлы (примеры) ---

class InputNode(BaseNode if NodeGraphQt else object):
    """
    Базовый узел для ввода данных (например, из ячейки Excel).
    """
    # Уникальный идентификатор узла (обычно определяется в NodeGraphQt)
    __identifier__ = 'nodes.input'
    # Имя узла, отображаемое в редакторе
    NODE_NAME = 'Input Data'

    def __init__(self):
        super(InputNode, self).__init__()
        # Добавляем выходной порт
        self.add_output('Output')
        # Можно добавить свойства (properties) для настройки (например, адрес ячейки)
        # self.create_property('cell_address', 'A1')


class OutputNode(BaseNode if NodeGraphQt else object):
    """
    Базовый узел для вывода данных (например, в ячейку Excel или в лог).
    """
    __identifier__ = 'nodes.output'
    NODE_NAME = 'Output Data'

    def __init__(self):
        super(OutputNode, self).__init__()
        # Добавляем входной порт
        self.add_input('Input')
        # Можно добавить свойства для настройки (например, адрес ячейки назначения)
        # self.create_property('target_cell', '')


class FormulaNode(BaseNode if NodeGraphQt else object):
    """
    Базовый узел для выполнения формул или скриптов.
    """
    __identifier__ = 'nodes.logic'
    NODE_NAME = 'Formula/Script'

    def __init__(self):
        super(FormulaNode, self).__init__()
        # Пример: вход для данных
        self.add_input('Data In')
        # Пример: выход для результата
        self.add_output('Result')
        # Свойство для текста формулы/скрипта
        # self.create_property('script', '=SUM(A1:A10)')


class NodeEditorWidget(QWidget):
    """
    Виджет нодового редактора, инкапсулирующий NodeGraphQt.
    Предназначен для встраивания в QDockWidget.
    """
    # Сигналы для взаимодействия с AppController/GUIController
    nodesConnected = Signal(object, object) # node1, node2
    nodeDoubleClicked = Signal(object) # node
    graphChanged = Signal() # Сигнал при любом изменении в графе

    def __init__(self, parent=None):
        """
        Инициализирует виджет нодового редактора.

        Args:
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self._setup_ui()
        self._setup_graph()

    def _setup_ui(self):
        """
        Настраивает пользовательский интерфейс виджета.
        """
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0) # Убираем отступы

        if NodeGraphQt is None:
            # Отображаем сообщение об ошибке, если NodeGraphQt не импортировался
            error_label = QMessageBox(self)
            error_label.setIcon(QMessageBox.Critical)
            error_label.setText("Ошибка загрузки NodeGraphQt")
            error_label.setInformativeText("Библиотека NodeGraphQt не найдена или несовместима. Убедитесь, что она установлена и совместима с PySide6.")
            error_label.setStandardButtons(QMessageBox.Ok)
            layout.addWidget(error_label)
            logger.error("NodeGraphQt не доступен. Виджет нодового редактора не функционален.")
            self.graph = None
            self.node_graph_widget = None
        else:
            # Создаем экземпляр NodeGraph
            self.graph = NodeGraph()
            
            # --- ИСПРАВЛЕНИЕ ---
            # NodeGraphQt.NodeGraph предоставляет метод viewer(), 
            # который возвращает QGraphicsView (подкласс QWidget), пригодный для addWidget.
            # Проверяем его наличие и тип.
            viewer = getattr(self.graph, 'viewer', None)
            if viewer and callable(viewer):
                self.node_graph_widget = viewer()
                logger.debug(f"Получен viewer из NodeGraph: {type(self.node_graph_widget)}")
            else:
                # Альтернативный способ: сам NodeGraph может быть QGraphicsView
                # Проверим, является ли self.graph QGraphicsView или QWidget
                if isinstance(self.graph, QWidget):
                     self.node_graph_widget = self.graph
                     logger.debug("NodeGraph напрямую является QWidget.")
                else:
                     self.node_graph_widget = None
                     logger.error("Не удалось получить QWidget из NodeGraph. viewer() не найден или не вызываем.")
            
            # Проверяем, что виджет получен и является QWidget
            if self.node_graph_widget and isinstance(self.node_graph_widget, QWidget):
                layout.addWidget(self.node_graph_widget)
                logger.debug("Виджет NodeGraphQt (QGraphicsView) добавлен в layout NodeEditorWidget.")
            else:
                # Отображаем сообщение об ошибке, если виджет не получен
                error_msg = "Не удалось получить виджет NodeGraphQt для отображения."
                logger.error(error_msg)
                error_label = QMessageBox(self)
                error_label.setIcon(QMessageBox.Critical)
                error_label.setText("Ошибка инициализации NodeGraphQt")
                error_label.setInformativeText(error_msg)
                error_label.setStandardButtons(QMessageBox.Ok)
                layout.addWidget(error_label)
            # --- КОНЕЦ ИСПРАВЛЕНИЯ ---

    def _setup_graph(self):
        """
        Настраивает граф, регистрирует узлы и подключает сигналы.
        """
        if self.graph is None:
            logger.warning("Граф не инициализирован, настройка пропущена.")
            return

        try:
            # Регистрируем пользовательские узлы
            self.graph.register_node(InputNode)
            self.graph.register_node(OutputNode)
            self.graph.register_node(FormulaNode)
            # Добавьте сюда регистрацию других пользовательских узлов

            # Подключаем сигналы NodeGraphQt к сигналам виджета
            # ВАЖНО: Сигнатуры сигналов NodeGraphQt могут отличаться.
            # Проверьте документацию NodeGraphQt для точных названий и аргументов сигналов.
            # Примеры подключения (сигнатуры могут быть неточны):
            # self.graph.node_double_clicked.connect(self._on_node_double_clicked)
            # self.graph.port_connected.connect(self._on_port_connected)
            # self.graph.graph_changed.connect(self.graphChanged.emit) # или аналогичный сигнал
            # self.graph.nodes_moved.connect(self.graphChanged.emit) # или аналогичный сигнал

            logger.info("Граф нодового редактора настроен и узлы зарегистрированы.")
        except Exception as e:
            logger.error(f"Ошибка при настройке графа NodeGraphQt: {e}", exc_info=True)

    # --- Методы для взаимодействия с графом ---

    def get_current_graph_data(self) -> Optional[Dict[str, Any]]:
        """
        Получает данные текущего графа в формате, пригодном для сериализации/сохранения.

        Returns:
            Optional[Dict[str, Any]]: Данные графа или None, если граф не инициализирован.
        """
        if self.graph is None:
            logger.warning("Граф не инициализирован, данные получить нельзя.")
            return None
        try:
            # NodeGraphQt обычно имеет метод для сериализации
            # Проверьте документацию NodeGraphQt для точного названия метода.
            # graph_data = self.graph.serialize_to_str() # Пример, может быть другой метод
            # или
            graph_data = self.graph.dump() # Часто используемый метод
            logger.debug("Данные графа получены для сериализации.")
            return graph_data
        except Exception as e:
            logger.error(f"Ошибка при получении данных графа: {e}", exc_info=True)
            return None

    def load_graph_data(self, data: Dict[str, Any]) -> bool:
        """
        Загружает данные графа из сериализованного формата.

        Args:
            data (Dict[str, Any]): Данные графа.

        Returns:
            bool: True, если загрузка успешна, иначе False.
        """
        if self.graph is None:
            logger.warning("Граф не инициализирован, данные загрузить нельзя.")
            return False
        try:
            # NodeGraphQt обычно имеет метод для десериализации
            # Проверьте документацию NodeGraphQt для точного названия метода.
            # self.graph.deserialize_from_str(data) # Пример
            # или
            self.graph.load_data(data) # Часто используемый метод
            logger.info("Данные графа успешно загружены.")
            self.graphChanged.emit() # Уведомляем о изменении
            return True
        except Exception as e:
            logger.error(f"Ошибка при загрузке данных графа: {e}", exc_info=True)
            return False

    def clear_graph(self):
        """
        Очищает текущий граф.
        """
        if self.graph is None:
            logger.warning("Граф не инициализирован, очистка невозможна.")
            return
        try:
            self.graph.clear_session()
            logger.info("Граф нодового редактора очищен.")
            self.graphChanged.emit() # Уведомляем о изменении
        except Exception as e:
            logger.error(f"Ошибка при очистке графа: {e}", exc_info=True)

    # --- Обработчики сигналов NodeGraphQt (примеры) ---
    # ВАЖНО: Сигнатуры зависят от NodeGraphQt. Ниже приведены примеры.
    # def _on_node_double_clicked(self, node):
    #     """Обработчик двойного клика по узлу."""
    #     logger.debug(f"Двойной клик по узлу: {node.NODE_NAME} ({node.id})")
    #     self.nodeDoubleClicked.emit(node)
    #
    # def _on_port_connected(self, input_port, output_port):
    #     """Обработчик соединения портов."""
    #     # input_port и output_port имеют типы Port из NodeGraphQt
    #     node1 = input_port.node() if input_port else None
    #     node2 = output_port.node() if output_port else None
    #     logger.debug(f"Соединены порты: {node1.NODE_NAME if node1 else 'None'} -> {node2.NODE_NAME if node2 else 'None'}")
    #     self.nodesConnected.emit(node1, node2)

    # --- Конец обработчиков ---
