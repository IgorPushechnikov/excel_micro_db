# backend/constructor/widgets/new_gui/import_mode_selector_new.py
"""
Виджет для выбора типа и режима импорта данных Excel.
"""

import logging
from typing import Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QButtonGroup,
    QRadioButton, QLabel
)
from PySide6.QtCore import Qt, Signal

from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ImportModeSelector(QWidget):
    """
    Виджет для выбора типа данных и режима импорта.
    """

    # Сигналы для уведомления о выборе
    type_selected = Signal(str)   # (selected_type)
    mode_selected = Signal(str)   # (selected_mode)

    # --- Константы для типов и режимов ---
    IMPORT_TYPES = {
        'all_data': "Все данные",
        'raw_data': "Сырые данные",
        'styles': "Стили",
        'charts': "Диаграммы",
        'formulas': "Формулы",
        'raw_data_pandas': "Сырые данные (pandas)"
    }

    IMPORT_MODES = {
        'all': "Всё",
        'selective': "Выборочно",
        'chunks': "Частями",
        'fast': "Быстрый (pandas)"
    }
    # ------------------------------------

    def __init__(self, parent=None):
        """
        Инициализирует селектор режима импорта.

        Args:
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self.setObjectName("ImportModeSelector")

        # --- Атрибуты виджета ---
        # Группы кнопок для типов и режимов
        self.type_button_group: Optional[QButtonGroup] = None
        self.mode_button_group: Optional[QButtonGroup] = None
        
        # Словари для хранения ссылок на радио-кнопки
        self.type_radio_buttons: dict[str, QRadioButton] = {}
        self.mode_radio_buttons: dict[str, QRadioButton] = {}
        # -----------------------

        self._setup_ui()
        self._setup_connections()
        
        # Установим начальные значения по умолчанию
        self._set_default_selection()

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # --- Группа для выбора типа данных ---
        type_group_box = QGroupBox("Тип данных для импорта")
        type_layout = QVBoxLayout(type_group_box)

        self.type_button_group = QButtonGroup(self)
        self.type_button_group.setExclusive(True)

        for type_key, type_label in self.IMPORT_TYPES.items():
            radio_button = QRadioButton(type_label, self)
            radio_button.setProperty("type_key", type_key) # Сохраняем ключ типа
            self.type_radio_buttons[type_key] = radio_button
            self.type_button_group.addButton(radio_button)
            type_layout.addWidget(radio_button)

        main_layout.addWidget(type_group_box)
        # ------------------------------------

        # --- Группа для выбора режима импорта ---
        mode_group_box = QGroupBox("Режим импорта")
        mode_layout = QVBoxLayout(mode_group_box)

        self.mode_button_group = QButtonGroup(self)
        self.mode_button_group.setExclusive(True)

        for mode_key, mode_label in self.IMPORT_MODES.items():
            radio_button = QRadioButton(mode_label, self)
            radio_button.setProperty("mode_key", mode_key) # Сохраняем ключ режима
            self.mode_radio_buttons[mode_key] = radio_button
            self.mode_button_group.addButton(radio_button)
            mode_layout.addWidget(radio_button)
            
        # --- НОВОЕ: Управление доступностью режимов ---
        # По умолчанию режим "Быстрый" доступен только для типа "Сырые данные (pandas)"
        self.mode_radio_buttons['fast'].setEnabled(False)
        # ---------------------------------------------

        main_layout.addWidget(mode_group_box)
        # ------------------------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        # Подключение сигналов групп кнопок
        self.type_button_group.buttonClicked.connect(self._on_type_button_clicked)
        self.mode_button_group.buttonClicked.connect(self._on_mode_button_clicked)

    def _set_default_selection(self):
        """Устанавливает начальные выбранные значения по умолчанию."""
        # По умолчанию выбираем "Все данные" и "Всё"
        if 'all_data' in self.type_radio_buttons:
            self.type_radio_buttons['all_data'].setChecked(True)
        if 'all' in self.mode_radio_buttons:
            self.mode_radio_buttons['all'].setChecked(True)
            
        # Обновляем доступность режимов для начального типа
        self._update_mode_availability('all_data')

    # --- Обработчики событий UI ---
    def _on_type_button_clicked(self, button: QRadioButton):
        """Обработчик клика по радио-кнопке типа."""
        type_key = button.property("type_key")
        if type_key:
            logger.debug(f"Выбран тип импорта: {type_key}")
            self.type_selected.emit(type_key)
            # Обновляем доступность режимов в зависимости от выбранного типа
            self._update_mode_availability(type_key)

    def _on_mode_button_clicked(self, button: QRadioButton):
        """Обработчик клика по радио-кнопке режима."""
        mode_key = button.property("mode_key")
        if mode_key:
            logger.debug(f"Выбран режим импорта: {mode_key}")
            self.mode_selected.emit(mode_key)
    # -----------------------------

    # --- Логика обновления доступности ---
    def _update_mode_availability(self, selected_type: str):
        """
        Обновляет доступность режимов импорта в зависимости от выбранного типа.
        
        Args:
            selected_type (str): Ключ выбранного типа импорта.
        """
        logger.debug(f"Обновление доступности режимов для типа: {selected_type}")
        
        # Режим "Быстрый" доступен только для типа "Сырые данные (pandas)"
        fast_mode_rb = self.mode_radio_buttons.get('fast')
        if fast_mode_rb:
            is_fast_available = (selected_type == 'raw_data_pandas')
            fast_mode_rb.setEnabled(is_fast_available)
            logger.debug(f"Режим 'Быстрый' {'доступен' if is_fast_available else 'недоступен'} для типа {selected_type}")
            
            # Если режим "Быстрый" был выбран, но стал недоступным, сбрасываем на "Всё"
            if not is_fast_available and fast_mode_rb.isChecked():
                all_mode_rb = self.mode_radio_buttons.get('all')
                if all_mode_rb:
                    all_mode_rb.setChecked(True)
                    self.mode_selected.emit('all')
                    logger.debug("Режим 'Быстрый' сброшен на 'Всё' из-за недоступности.")
    # ------------------------------------

    # --- Методы для получения выбранных значений ---
    def get_selected_type(self) -> str:
        """
        Возвращает ключ выбранного типа импорта.
        
        Returns:
            str: Ключ выбранного типа (например, 'all_data').
        """
        checked_button = self.type_button_group.checkedButton()
        if checked_button:
            type_key = checked_button.property("type_key")
            if type_key:
                return type_key
        # Если ничего не выбрано, возвращаем значение по умолчанию
        return 'all_data'

    def get_selected_mode(self) -> str:
        """
        Возвращает ключ выбранного режима импорта.
        
        Returns:
            str: Ключ выбранного режима (например, 'all').
        """
        checked_button = self.mode_button_group.checkedButton()
        if checked_button:
            mode_key = checked_button.property("mode_key")
            if mode_key:
                return mode_key
        # Если ничего не выбрано, возвращаем значение по умолчанию
        return 'all'
    # ------------------------------------------------

    # --- Методы для программного выбора (опционально) ---
    def set_selected_type(self, type_key: str):
        """
        Программно выбирает тип импорта.
        
        Args:
            type_key (str): Ключ типа для выбора.
        """
        radio_button = self.type_radio_buttons.get(type_key)
        if radio_button:
            radio_button.setChecked(True)
            # Вручную вызываем обработчик, так как setChecked не генерирует сигнал
            self._on_type_button_clicked(radio_button)

    def set_selected_mode(self, mode_key: str):
        """
        Программно выбирает режим импорта.
        
        Args:
            mode_key (str): Ключ режима для выбора.
        """
        radio_button = self.mode_radio_buttons.get(mode_key)
        if radio_button:
            radio_button.setChecked(True)
            # Вручную вызываем обработчик, так как setChecked не генерирует сигнал
            self._on_mode_button_clicked(radio_button)
    # ----------------------------------------------------
