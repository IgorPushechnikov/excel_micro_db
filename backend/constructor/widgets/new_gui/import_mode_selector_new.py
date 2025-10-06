# backend/constructor/widgets/new_gui/import_mode_selector_new.py
"""
Упрощённый виджет для выбора режима импорта данных Excel.
Все режимы представлены в одном списке радиокнопок.
"""

import logging
from typing import Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QButtonGroup, QRadioButton
)
from PySide6.QtCore import Qt, Signal

from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ImportModeSelector(QGroupBox):
    """
    Виджет для выбора одного из 7 упрощённых режимов импорта.
    """

    # Сигнал для уведомления о выборе режима
    mode_selected = Signal(str)  # (selected_mode_key)

    # --- Константы для режимов ---
    # Ключи режимов
    MODE_KEYS = [
        'all_openpyxl',
        'raw_openpyxl',
        'styles_openpyxl',
        'charts_openpyxl',
        'formulas_openpyxl',
        # 'raw_fast_pandas', # <-- УДАЛЁН: Режим больше не поддерживается
        'chunks_openpyxl',
        'auto'  # <-- НОВЫЙ РЕЖИМ
    ]

    # Метки для отображения
    MODE_LABELS = [
        "Всё - openpyxl",
        "Только данные - openpyxl",
        "Стили - openpyxl",
        "Диаграммы - openpyxl",
        "Формулы - openpyxl",
        # "Быстрый только данные - pandas", # <-- УДАЛЕНА: Метка для удалённого режима
        "Частями - openpyxl (Экспериментальный)",
        "Авто (Данные-Pandas, Стили/OpenPyxl, Диаграммы/OpenPyxl, Формулы/OpenPyxl)" # <-- НОВАЯ МЕТКА
    ]

    # Словарь для удобства получения метки по ключу (если понадобится)
    MODE_KEY_TO_LABEL = dict(zip(MODE_KEYS, MODE_LABELS))
    # ------------------------------------

    def __init__(self, parent=None):
        """
        Инициализирует селектор режима импорта.

        Args:
            parent: Родительский объект Qt.
        """
        super().__init__("Режим импорта", parent)
        self.setObjectName("ImportModeSelector")

        # --- Атрибуты виджета ---
        # Группа кнопок для режимов
        self.mode_button_group: Optional[QButtonGroup] = None
        # Словарь для хранения ссылок на радио-кнопки
        self.mode_radio_buttons: dict[str, QRadioButton] = {}
        # -----------------------

        self._setup_ui()
        self._setup_connections()
        # Установим начальное значение по умолчанию
        self._set_default_selection()

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        main_layout = QVBoxLayout(self)

        # --- Группа для выбора РЕЖИМА (вертикально) ---
        self.mode_button_group = QButtonGroup(self)
        self.mode_button_group.setExclusive(True)  # Только одна кнопка может быть выбрана

        for mode_key, mode_label in zip(self.MODE_KEYS, self.MODE_LABELS):
            radio_button = QRadioButton(mode_label, self)
            radio_button.setProperty("mode_key", mode_key)
            self.mode_radio_buttons[mode_key] = radio_button
            self.mode_button_group.addButton(radio_button)
            main_layout.addWidget(radio_button)
        # ---------------------------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        # Подключение сигнала группы кнопок РЕЖИМОВ
        assert self.mode_button_group is not None
        self.mode_button_group.buttonClicked.connect(self._on_mode_button_clicked)

    def _set_default_selection(self):
        """Устанавливает начальный выбранный режим по умолчанию."""
        # По умолчанию выбираем "Всё - openpyxl"
        default_mode_key = 'all_openpyxl'
        radio_button = self.mode_radio_buttons.get(default_mode_key)
        if radio_button:
            radio_button.setChecked(True)

    # --- Обработчики событий UI ---
    def _on_mode_button_clicked(self, button: QRadioButton):
        """Обработчик клика по радио-кнопке режима."""
        mode_key = button.property("mode_key")
        if mode_key:
            logger.debug(f"Выбран режим импорта: {mode_key} ({self.MODE_KEY_TO_LABEL.get(mode_key, 'Неизвестный')})")
            self.mode_selected.emit(mode_key)
    # -----------------------------

    # --- Методы для получения выбранных значений ---
    def get_selected_mode(self) -> str:
        """
        Возвращает ключ выбранного режима импорта.

        Returns:
            str: Ключ выбранного режима (например, 'all_openpyxl').
        """
        assert self.mode_button_group is not None
        checked_button = self.mode_button_group.checkedButton()
        if checked_button:
            mode_key = checked_button.property("mode_key")
            if mode_key:
                return mode_key
        # Если ничего не выбрано (хотя по умолчанию должно быть), возвращаем значение по умолчанию
        return 'all_openpyxl'
    # ------------------------------------------------