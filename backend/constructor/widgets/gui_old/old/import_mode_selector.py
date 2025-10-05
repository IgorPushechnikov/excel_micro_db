# backend/constructor/widgets/new_gui/import_mode_selector.py
"""
Виджет для выбора типа и режима импорта в диалоге импорта.
"""

import logging
from typing import Optional, Dict, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QHBoxLayout, QRadioButton, 
    QLabel, QButtonGroup, QSizePolicy
)
from PySide6.QtCore import Qt, Signal

from backend.utils.logger import get_logger

logger = get_logger(__name__)

class ImportModeSelector(QWidget):
    """
    Виджет для выбора типа данных и режима импорта.
    """

    # Сигнал, испускаемый при изменении выбора
    selection_changed = Signal(str, str) # (import_type, import_mode)

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
    # -------------------------------------

    def __init__(self, parent=None):
        """
        Инициализирует виджет выбора импорта.

        Args:
            parent: Родительский объект Qt.
        """
        super().__init__(parent)
        self._selected_import_type: str = 'all_data'
        self._selected_import_mode: str = 'all'
        
        # --- Группы кнопок для управления exclusivity ---
        self._type_button_group: Optional[QButtonGroup] = None
        self._mode_button_group: Optional[QButtonGroup] = None
        # --------------------------------------------------
        
        # --- Атрибуты для радиокнопок ---
        # Типы данных
        self.all_data_radio: Optional[QRadioButton] = None
        self.raw_data_radio: Optional[QRadioButton] = None
        self.styles_radio: Optional[QRadioButton] = None
        self.charts_radio: Optional[QRadioButton] = None
        self.formulas_radio: Optional[QRadioButton] = None
        self.raw_data_pandas_radio: Optional[QRadioButton] = None
        
        # Режимы импорта
        self.mode_all_radio: Optional[QRadioButton] = None
        self.mode_selective_radio: Optional[QRadioButton] = None
        self.mode_chunks_radio: Optional[QRadioButton] = None
        self.mode_fast_radio: Optional[QRadioButton] = None
        # ---------------------------------
        
        self._setup_ui()
        self._setup_connections()
        
        # Устанавливаем начальное состояние
        self._update_mode_availability()

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # --- Группа для типа данных ---
        type_group_box = QGroupBox("Тип данных", self)
        type_layout = QVBoxLayout(type_group_box)
        
        self._type_button_group = QButtonGroup(self)
        self._type_button_group.setExclusive(True)

        self.all_data_radio = QRadioButton(self.IMPORT_TYPES['all_data'], self)
        self.all_data_radio.setChecked(True)
        self._type_button_group.addButton(self.all_data_radio)
        type_layout.addWidget(self.all_data_radio)

        self.raw_data_radio = QRadioButton(self.IMPORT_TYPES['raw_data'], self)
        self._type_button_group.addButton(self.raw_data_radio)
        type_layout.addWidget(self.raw_data_radio)

        self.styles_radio = QRadioButton(self.IMPORT_TYPES['styles'], self)
        self._type_button_group.addButton(self.styles_radio)
        type_layout.addWidget(self.styles_radio)

        self.charts_radio = QRadioButton(self.IMPORT_TYPES['charts'], self)
        self._type_button_group.addButton(self.charts_radio)
        type_layout.addWidget(self.charts_radio)

        self.formulas_radio = QRadioButton(self.IMPORT_TYPES['formulas'], self)
        self._type_button_group.addButton(self.formulas_radio)
        type_layout.addWidget(self.formulas_radio)

        self.raw_data_pandas_radio = QRadioButton(self.IMPORT_TYPES['raw_data_pandas'], self)
        self._type_button_group.addButton(self.raw_data_pandas_radio)
        type_layout.addWidget(self.raw_data_pandas_radio)

        main_layout.addWidget(type_group_box)
        # -----------------------------

        # --- Группа для режима импорта ---
        mode_group_box = QGroupBox("Режим импорта", self)
        mode_layout = QVBoxLayout(mode_group_box)
        
        self._mode_button_group = QButtonGroup(self)
        self._mode_button_group.setExclusive(True)

        self.mode_all_radio = QRadioButton(self.IMPORT_MODES['all'], self)
        self.mode_all_radio.setChecked(True)
        self._mode_button_group.addButton(self.mode_all_radio)
        mode_layout.addWidget(self.mode_all_radio)

        self.mode_selective_radio = QRadioButton(self.IMPORT_MODES['selective'], self)
        self._mode_button_group.addButton(self.mode_selective_radio)
        mode_layout.addWidget(self.mode_selective_radio)

        self.mode_chunks_radio = QRadioButton(self.IMPORT_MODES['chunks'], self)
        self._mode_button_group.addButton(self.mode_chunks_radio)
        mode_layout.addWidget(self.mode_chunks_radio)

        self.mode_fast_radio = QRadioButton(self.IMPORT_MODES['fast'], self)
        self.mode_fast_radio.setEnabled(False) # По умолчанию недоступен
        self._mode_button_group.addButton(self.mode_fast_radio)
        mode_layout.addWidget(self.mode_fast_radio)

        main_layout.addWidget(mode_group_box)
        # ---------------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        # --- Подключение сигналов изменения выбора типа ---
        self.all_data_radio.toggled.connect(self._on_import_type_toggled)
        self.raw_data_radio.toggled.connect(self._on_import_type_toggled)
        self.styles_radio.toggled.connect(self._on_import_type_toggled)
        self.charts_radio.toggled.connect(self._on_import_type_toggled)
        self.formulas_radio.toggled.connect(self._on_import_type_toggled)
        self.raw_data_pandas_radio.toggled.connect(self._on_import_type_toggled)
        # --------------------------------------------------

        # --- Подключение сигналов изменения выбора режима ---
        self.mode_all_radio.toggled.connect(self._on_import_mode_toggled)
        self.mode_selective_radio.toggled.connect(self._on_import_mode_toggled)
        self.mode_chunks_radio.toggled.connect(self._on_import_mode_toggled)
        self.mode_fast_radio.toggled.connect(self._on_import_mode_toggled)
        # ---------------------------------------------------

    # --- Обработчики событий ---
    def _on_import_type_toggled(self, checked: bool):
        """
        Обработчик переключения радиокнопок типа импорта.
        """
        if not checked:
            return # Обрабатываем только включение
            
        sender = self.sender()
        if sender == self.all_data_radio:
            self._selected_import_type = 'all_data'
        elif sender == self.raw_data_radio:
            self._selected_import_type = 'raw_data'
        elif sender == self.styles_radio:
            self._selected_import_type = 'styles'
        elif sender == self.charts_radio:
            self._selected_import_type = 'charts'
        elif sender == self.formulas_radio:
            self._selected_import_type = 'formulas'
        elif sender == self.raw_data_pandas_radio:
            self._selected_import_type = 'raw_data_pandas'
        else:
            return # Неизвестный отправитель
            
        logger.debug(f"Выбран тип импорта: {self._selected_import_type}")
        self._update_mode_availability()
        self.selection_changed.emit(self._selected_import_type, self._selected_import_mode)

    def _on_import_mode_toggled(self, checked: bool):
        """
        Обработчик переключения радиокнопок режима импорта.
        """
        if not checked:
            return # Обрабатываем только включение
            
        sender = self.sender()
        if sender == self.mode_all_radio:
            self._selected_import_mode = 'all'
        elif sender == self.mode_selective_radio:
            self._selected_import_mode = 'selective'
        elif sender == self.mode_chunks_radio:
            self._selected_import_mode = 'chunks'
        elif sender == self.mode_fast_radio:
            self._selected_import_mode = 'fast'
        else:
            return # Неизвестный отправитель
            
        logger.debug(f"Выбран режим импорта: {self._selected_import_mode}")
        # Проверка на допустимость комбинации уже сделана в _update_mode_availability
        self.selection_changed.emit(self._selected_import_type, self._selected_import_mode)
    # --------------------------

    def _update_mode_availability(self):
        """
        Обновляет доступность режимов импорта в зависимости от выбранного типа.
        """
        is_pandas_type = self._selected_import_type == 'raw_data_pandas'
        
        # Режим "Быстрый (pandas)" доступен только для типа "Сырые данные (pandas)"
        self.mode_fast_radio.setEnabled(is_pandas_type)
        
        # Если выбран тип "Сырые данные (pandas)", автоматически выбираем режим "Быстрый"
        if is_pandas_type and self.mode_fast_radio.isEnabled():
            self.mode_fast_radio.setChecked(True)
            self._selected_import_mode = 'fast'
            logger.debug("Автоматически выбран режим 'Быстрый (pandas)' для типа 'Сырые данные (pandas)'.")
        # Если был выбран режим "Быстрый", но тип изменили не на "Сырые данные (pandas)",
        # сбрасываем на режим "Всё"
        elif not is_pandas_type and self._selected_import_mode == 'fast':
            self.mode_all_radio.setChecked(True)
            self._selected_import_mode = 'all'
            logger.debug("Сброшен режим на 'Всё', так как тип не 'Сырые данные (pandas)'.")
            
        logger.debug(f"Доступность режимов обновлена. Текущий тип: {self._selected_import_type}, режим: {self._selected_import_mode}")

    # --- Публичные методы для получения выбора ---
    def get_selected_import_type(self) -> str:
        """
        Возвращает выбранный тип импорта.

        Returns:
            str: Выбранный тип импорта (например, 'all_data').
        """
        return self._selected_import_type

    def get_selected_import_mode(self) -> str:
        """
        Возвращает выбранный режим импорта.

        Returns:
            str: Выбранный режим импорта (например, 'all').
        """
        return self._selected_import_mode
        
    def set_selection(self, import_type: str, import_mode: str):
        """
        Устанавливает выбранные тип и режим импорта программно.

        Args:
            import_type (str): Тип импорта.
            import_mode (str): Режим импорта.
        """
        # Устанавливаем тип
        if import_type == 'all_data':
            self.all_data_radio.setChecked(True)
        elif import_type == 'raw_data':
            self.raw_data_radio.setChecked(True)
        elif import_type == 'styles':
            self.styles_radio.setChecked(True)
        elif import_type == 'charts':
            self.charts_radio.setChecked(True)
        elif import_type == 'formulas':
            self.formulas_radio.setChecked(True)
        elif import_type == 'raw_data_pandas':
            self.raw_data_pandas_radio.setChecked(True)
        else:
            logger.warning(f"Неизвестный тип импорта для установки: {import_type}")
            
        # Устанавливаем режим (проверка допустимости будет в _update_mode_availability)
        if import_mode == 'all':
            self.mode_all_radio.setChecked(True)
        elif import_mode == 'selective':
            self.mode_selective_radio.setChecked(True)
        elif import_mode == 'chunks':
            self.mode_chunks_radio.setChecked(True)
        elif import_mode == 'fast':
            self.mode_fast_radio.setChecked(True)
        else:
            logger.warning(f"Неизвестный режим импорта для установки: {import_mode}")
            
        # Обновляем доступность и, возможно, скорректируем выбор
        self._update_mode_availability()
        
        # Испускаем сигнал, если выбор действительно изменился
        if (self._selected_import_type, self._selected_import_mode) != (import_type, import_mode):
             self.selection_changed.emit(self._selected_import_type, self._selected_import_mode)
