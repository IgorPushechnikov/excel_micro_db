# backend/constructor/widgets/new_gui/import_mode_selector_new.py
"""
Виджет для выбора типа и режима импорта данных Excel.
Режимы отображаются горизонтально над типами.
При выборе режима показываются/скрываются опции и комментарии.
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

class ImportModeSelector(QGroupBox):
    """
    Виджет для выбора типа данных и режима импорта.
    Режимы - сверху, горизонтально.
    Типы - снизу, вертикально.
    Опции - под типами, динамически добавляются/удаляются в макет в зависимости от режима.
    Комментарий - под опциями, скрывается/показывается и обновляется в зависимости от режима.
    """

    # Сигналы для уведомления о выборе
    type_selected = Signal(str)   # (selected_type)
    mode_selected = Signal(str)   # (selected_mode)

    # --- Константы для типов и режимов ---
    # Типы данных для импорта
    IMPORT_TYPES = {
        'all_data': "Все данные",
        'raw_data': "Сырые данные",
        'styles': "Стили",
        'charts': "Диаграммы",
        'formulas': "Формулы",
        'raw_data_pandas': "Сырые данные (pandas)" # Пока остается, но может быть скрыт в "Выборочно"
    }

    # Режимы импорта
    IMPORT_MODES = {
        'all': "Всё",
        'selective': "Выборочно",
        'chunks': "Частями",
        'fast': "Быстрый (pandas)"
    }

    # Опции, которые могут отображаться внутри режима "Выборочно"
    SELECTIVE_OPTIONS = {
        'raw_data': "Сырые данные",
        'styles': "Стили",
        'charts': "Диаграммы",
        'formulas': "Формулы"
    }

    # Комментарии для режимов
    MODE_COMMENTS = {
        'all': "Для лёгких файлов. Применяется openpyxl.",
        'fast': "Для работы с тяжёлыми файлами, только данные.",
        'chunks': "Экспериментальный режим."
        # 'selective' не имеет комментария по умолчанию
    }
    # ------------------------------------

    def __init__(self, parent=None):
        """
        Инициализирует селектор режима импорта.

        Args:
            parent: Родительский объект Qt.
        """
        super().__init__("Тип и режим импорта", parent) # Группа теперь охватывает всё
        self.setObjectName("ImportModeSelector")

        # --- Атрибуты виджета ---
        # Группы кнопок для режимов, типов и опций
        self.mode_button_group: Optional[QButtonGroup] = None
        self.type_button_group: Optional[QButtonGroup] = None
        self.option_button_group: Optional[QButtonGroup] = None # Для опций "Выборочно"
        # Словари для хранения ссылок на радио-кнопки
        self.mode_radio_buttons: dict[str, QRadioButton] = {}
        self.type_radio_buttons: dict[str, QRadioButton] = {}
        self.option_radio_buttons: dict[str, QRadioButton] = {} # Для опций "Выборочно"
        # Виджеты для опций и комментария
        self.options_group_widget: Optional[QWidget] = None
        self.comment_label: Optional[QLabel] = None
        # Ссылка на основной макет для динамического добавления/удаления
        self.main_layout: Optional[QVBoxLayout] = None
        # -----------------------

        self._setup_ui()
        self._setup_connections()
        # Установим начальные значения по умолчанию
        self._set_default_selection()

    def _setup_ui(self):
        """Создаёт элементы интерфейса."""
        self.main_layout = QVBoxLayout(self) # Сохраняем ссылку
        self.main_layout.setContentsMargins(0, 0, 0, 0)

        # --- Группа для выбора РЕЖИМА (горизонтально) ---
        mode_group_box = QGroupBox(title="Режим импорта", parent=self)
        mode_layout = QHBoxLayout(mode_group_box) # Горизонтальный Layout

        self.mode_button_group = QButtonGroup(self)
        self.mode_button_group.setExclusive(True)

        for mode_key, mode_label in self.IMPORT_MODES.items():
            radio_button = QRadioButton(mode_label, self)
            radio_button.setProperty("mode_key", mode_key)
            self.mode_radio_buttons[mode_key] = radio_button
            self.mode_button_group.addButton(radio_button)
            mode_layout.addWidget(radio_button)

        self.main_layout.addWidget(mode_group_box)
        # ---------------------------------------------

        # --- Группа для выбора ТИПА (вертикально) ---
        type_group_box = QGroupBox(title="Тип данных для импорта", parent=self)
        type_layout = QVBoxLayout(type_group_box)

        self.type_button_group = QButtonGroup(self)
        self.type_button_group.setExclusive(True)

        for type_key, type_label in self.IMPORT_TYPES.items():
            radio_button = QRadioButton(type_label, self)
            radio_button.setProperty("type_key", type_key)
            self.type_radio_buttons[type_key] = radio_button
            self.type_button_group.addButton(radio_button)
            type_layout.addWidget(radio_button)

        self.main_layout.addWidget(type_group_box)
        # ------------------------------------------

        # --- Виджет для ОПЦИЙ (например, для "Выборочно") ---
        # Создаем, но НЕ добавляем в main_layout сразу.
        # Он будет добавляться/удаляться динамически.
        self.options_group_widget = QWidget(self)
        self.options_group_widget.setVisible(False)  # Явно скрываем при создании
        options_layout = QVBoxLayout(self.options_group_widget)

        # Добавляем только те опции, которые нужны для "Выборочно"
        self.option_button_group = QButtonGroup(self)
        self.option_button_group.setExclusive(False) # Можно выбрать несколько опций

        for opt_key, opt_label in self.SELECTIVE_OPTIONS.items():
            # Создадим отдельные QRadioButton для опций.
            # ИСПРАВЛЕНО: родитель теперь self.options_group_widget
            option_radio_button = QRadioButton(opt_label, self.options_group_widget)
            option_radio_button.setProperty("option_key", opt_key)
            self.option_radio_buttons[opt_key] = option_radio_button
            self.option_button_group.addButton(option_radio_button)
            options_layout.addWidget(option_radio_button)
            logger.debug(f"Добавлена опция выбора '{opt_key}' как отдельный элемент.")

        # main_layout.addWidget(self.options_group_widget) # <-- УДАЛЕНО
        # ---------------------------------------------

        # --- Метка для комментария ---
        self.comment_label = QLabel("", self)
        self.comment_label.setWordWrap(True) # Перенос строк
        self.comment_label.setVisible(False) # Скрыта по умолчанию
        self.main_layout.addWidget(self.comment_label)
        # -----------------------------

    def _setup_connections(self):
        """Подключает сигналы к слотам."""
        # Подключение сигналов группы кнопок РЕЖИМОВ
        assert self.mode_button_group is not None
        self.mode_button_group.buttonClicked.connect(self._on_mode_button_clicked)

        # Подключение сигналов группы кнопок ТИПОВ
        assert self.type_button_group is not None
        self.type_button_group.buttonClicked.connect(self._on_type_button_clicked)

        # Подключение сигналов группы кнопок ОПЦИЙ (если понадобится)
        # self.option_button_group.buttonClicked.connect(self._on_option_button_clicked) # Пример


    def _set_default_selection(self):
        """Устанавливает начальные выбранные значения по умолчанию."""
        # По умолчанию выбираем "Всё" и "Все данные"
        if 'all' in self.mode_radio_buttons:
            self.mode_radio_buttons['all'].setChecked(True)
        if 'all_data' in self.type_radio_buttons:
            self.type_radio_buttons['all_data'].setChecked(True)

        # Обновляем интерфейс для начального режима
        self._update_ui_for_mode('all')

    # --- Обработчики событий UI ---
    def _on_mode_button_clicked(self, button: QRadioButton):
        """Обработчик клика по радио-кнопке режима."""
        mode_key = button.property("mode_key")
        if mode_key:
            logger.debug(f"Выбран режим импорта: {mode_key}")
            self.mode_selected.emit(mode_key)
            # Обновляем UI в зависимости от выбранного режима
            self._update_ui_for_mode(mode_key)

    def _on_type_button_clicked(self, button: QRadioButton):
        """Обработчик клика по радио-кнопке типа."""
        type_key = button.property("type_key")
        if type_key:
            logger.debug(f"Выбран тип импорта: {type_key}")
            self.type_selected.emit(type_key)
            # В текущей логике тип не влияет на UI, но можно добавить
    # -----------------------------

    # --- Логика обновления UI ---
    def _update_ui_for_mode(self, selected_mode: str):
        """
        Обновляет макет (добавляет/удаляет опции) и видимость комментария
        в зависимости от выбранного режима.

        Args:
            selected_mode (str): Ключ выбранного режима импорта.
        """
        logger.debug(f"Обновление UI для режима: {selected_mode}")

        # Добавим assert для проверки, что необходимые атрибуты инициализированы
        assert self.options_group_widget is not None
        assert self.main_layout is not None
        assert self.comment_label is not None
        # Также убедимся, что они являются экземплярами QWidget (они такими и являются, но Pylance может не знать)
        assert isinstance(self.options_group_widget, QWidget)
        assert isinstance(self.comment_label, QWidget)

        # --- НОВАЯ ЛОГИКА: Динамическое управление макетом ---
        # 1. Управление виджетом опций
        if selected_mode == 'selective':
            # Найти индекс comment_label для вставки перед ним
            comment_index = self.main_layout.indexOf(self.comment_label)
            current_options_index = self.main_layout.indexOf(self.options_group_widget)

            # Если опции еще не добавлены или добавлены не перед comment_label
            if current_options_index == -1 or current_options_index >= comment_index:
                # Убедимся, что опции удалены, если были добавлены в другое место
                if current_options_index != -1:
                    self.main_layout.removeWidget(self.options_group_widget)

                # Вставляем перед comment_label
                if comment_index != -1:
                    self.main_layout.insertWidget(comment_index, self.options_group_widget)
                else:
                    # Если comment_label не найден (неожиданно), добавим в конец
                    self.main_layout.addWidget(self.options_group_widget)
                # Показываем виджет НЕЗАВИСИМО от способа добавления
                self.options_group_widget.setVisible(True)
                logger.debug("Виджет опций добавлен в макет.")
        else:
            # Удалить опции из макета, если они там есть
            current_options_index = self.main_layout.indexOf(self.options_group_widget)
            if current_options_index != -1:
                self.main_layout.removeWidget(self.options_group_widget)
                self.options_group_widget.setVisible(False)  # Скрываем виджет
                logger.debug("Виджет опций удален из макета.")
        # --- КОНЕЦ НОВОЙ ЛОГИКИ ---

        # 2. Управление комментарием
        self.comment_label.setVisible(False)
        if selected_mode in self.MODE_COMMENTS:
            self.comment_label.setText(self.MODE_COMMENTS[selected_mode])
            self.comment_label.setVisible(True)
            logger.debug(f"Показан комментарий для режима '{selected_mode}'.")

    # ------------------------------------

    # --- Методы для получения выбранных значений ---
    def get_selected_type(self) -> str:
        """
        Возвращает ключ выбранного типа импорта.

        Returns:
            str: Ключ выбранного типа (например, 'all_data').
        """
        assert self.type_button_group is not None
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
        assert self.mode_button_group is not None
        checked_button = self.mode_button_group.checkedButton()
        if checked_button:
            mode_key = checked_button.property("mode_key")
            if mode_key:
                return mode_key
        # Если ничего не выбрано, возвращаем значение по умолчанию
        return 'all'

    # --- Методы для получения выбранных опций (для режима "Выборочно") ---
    def get_selected_options(self) -> list[str]:
        """
        Возвращает список ключей выбранных опций (для режима "Выборочно").
        Возвращает пустой список, если текущий режим не 'selective'.

        Returns:
            list[str]: Список ключей выбранных опций (например, ['raw_data', 'styles']).
        """
        # Логическая защита: опции актуальны ТОЛЬКО в режиме 'selective'
        if self.get_selected_mode() != 'selective':
            return []
        
        selected_options = []
        if self.option_button_group:
            for button in self.option_button_group.buttons():
                if button.isChecked():
                    opt_key = button.property("option_key")
                    if opt_key:
                        selected_options.append(opt_key)
        return selected_options
    # ------------------------------------------------

    # --- Методы для программного выбора (опционально) ---
    def set_selected_mode(self, mode_key: str):
        """
        Программно выбирает режим импорта.
        Обновляет UI.

        Args:
            mode_key (str): Ключ режима для выбора.
        """
        radio_button = self.mode_radio_buttons.get(mode_key)
        if radio_button:
            radio_button.setChecked(True)
            # Вручную вызываем обработчик, так как setChecked не генерирует сигнал
            self._on_mode_button_clicked(radio_button)

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

    # Метод для программного выбора опций (для "Выборочно")
    def set_selected_options(self, option_keys: list[str]):
        """
        Программно выбирает опции импорта (для режима "Выборочно").

        Args:
            option_keys (list[str]): Список ключей опций для выбора.
        """
        if self.option_button_group:
            for button in self.option_button_group.buttons():
                opt_key = button.property("option_key")
                if opt_key:
                    button.setChecked(opt_key in option_keys)
    # ----------------------------------------------------