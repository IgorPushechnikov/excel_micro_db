# backend/constructor/widgets/new_gui/qt_model_adapter.py
"""
Модуль-адаптер для связи между AppController/БД и QAbstractTableModel (PySide6).
Предоставляет данные из БД для отображения в QTableView.
"""

import logging
from typing import Any, Dict, List, Optional
def _load_styles_from_controller(self):
        """
        Загружает стили ячеек из AppController.
        """
        # TODO: Реализовать загрузку стилей через AppController
        # AppController должен предоставить метод, например, get_sheet_styles(sheet_name)
        # и возвратить список словарей с range_address и style_attributes
        # Затем этот список нужно преобразовать в self._styles
        # Пока что оставлю как заглушку
        logger.info(f"Загрузка стилей для листа '{self.sheet_name}' (пока не реализовано).")
        # Примерный каркас:
        # styles_list = self.app_controller.get_sheet_styles(self.sheet_name)
        # for style_item in styles_list:
        #     range_addr = style_item['range_address']
        #     style_attrs = json.loads(style_item['style_attributes'])
        #     # Преобразовать range_addr в координаты (row_start, col_start, row_end, col_end)
        #     row_start, col_start, row_end, col_end = self._xl_range_to_coords(range_addr)
        #     # Заполнить self._styles для каждой ячейки в диапазоне
        #     for r in range(row_start, row_end + 1):
        #         for c in range(col_start, col_end + 1):
        #             self._styles[(r, c)] = style_attrs # Учитывать пересечения стилей
        # pass

    def _load_merged_cells_from_controller(self):
        """
        Загружает объединённые ячейки из AppController.
        """
        # TODO: Реализовать загрузку объединений через AppController
        # AppController должен предоставить метод, например, get_sheet_merged_cells(sheet_name)
        # и возвратить список строк адресов ('A1:B2')
        # Затем этот список нужно преобразовать в self._merged_cells
        # Пока что оставлю как заглушку
        logger.info(f"Загрузка объединений для листа '{self.sheet_name}' (пока не реализовано).")
        # Примерный каркас:
        # merged_ranges = self.app_controller.get_sheet_merged_cells(self.sheet_name)
        # for range_str in merged_ranges:
        #     row_start, col_start, row_end, col_end = self._xl_range_to_coords(range_str)
        #     self._merged_cells.append((row_start, col_start, row_end, col_end))
        # pass

    def _xl_cell_to_row_col(self, cell: str) -> tuple[int, int]:
        """
        Преобразует адрес ячейки Excel (e.g., 'A1') в индексы строки и столбца (0-based).
        """
        # Используем логику из xlsxwriter_exporter для согласованности
        col_str = ""
        row_str = ""
        for char in cell:
            if char.isalpha():
                col_str += char.upper()
            elif char.isdigit():
                row_str += char

        if not col_str or not row_str:
            error_msg = f"Неверный формат адреса ячейки: {cell}"
            logger.error(f"[КООРД] {error_msg}")
            raise ValueError(error_msg)

        row = int(row_str) - 1 # 0-based
        col = 0
        for c in col_str:
            col = col * 26 + (ord(c) - ord('A') + 1)
        col -= 1 # 0-based
        result = (row, col)
        # logger.debug(f"[КООРД] Результат для '{cell}': {result}")
        return result

    def _index_to_column_name(self, index: int) -> str:
        """
        Преобразует индекс столбца (0-based) в его буквенное обозначение Excel (A, B, ..., Z, AA, AB, ...).
        """
        if index < 0:
            return ""
        result = ""
        while index >= 0:
            result = chr(ord('A') + index % 26) + result
            index = index // 26 - 1
        return result

    def _xl_range_to_coords(self, range_str: str) -> tuple[int, int, int, int]:
        """
        Преобразует диапазон Excel (e.g., 'A1:B10') в координаты (row_start, col_start, row_end, col_end) (0-based).
        """
        # Используем логику из xlsxwriter_exporter для согласованности
        if ':' not in range_str:
            # Это одиночная ячейка
            # logger.debug(f"[КООРД] Диапазон '{range_str}' - это одиночная ячейка.")
            r, c = self._xl_cell_to_row_col(range_str)
            coords = (r, c, r, c)
            # logger.debug(f"[КООРД] Результат для '{range_str}': {coords}")
            return coords

        start_cell, end_cell = range_str.split(':', 1)
        # logger.debug(f"[КООРД] Разделение диапазона на '{start_cell}' и '{end_cell}'.")
        row_start, col_start = self._xl_cell_to_row_col(start_cell)
        row_end, col_end = self._xl_cell_to_row_col(end_cell)
        coords = (row_start, col_start, row_end, col_end)
        # logger.debug(f"[КООРД] Результат для диапазона '{range_str}': {coords}")
        return coords

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Возвращает количество строк."""
        if parent.isValid():
            return 0
        return len(self._data)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Возвращает количество столбцов."""
        if parent.isValid():
            return 0
        return len(self._data[0]) if self._data else 0

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        """Возвращает данные для указанной ячейки и роли."""
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if row >= len(self._data) or col >= len(self._data[0]):
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._data[row][col]
            # Возвращаем строковое представление значения
            return str(value) if value is not None else ""
        elif role == Qt.ItemDataRole.EditRole:
            # Для редактирования возвращаем "сырое" значение
            return self._data[row][col]
        elif role == Qt.ItemDataRole.BackgroundRole:
            # TODO: Вернуть QBrush для фона ячейки на основе стиля
            style = self._styles.get((row, col), {})
            bg_color_hex = style.get('bg_color')
            if bg_color_hex:
                return QBrush(QColor(f"#{bg_color_hex}"))
        elif role == Qt.ItemDataRole.ForegroundRole:
            # TODO: Вернуть QBrush для текста ячейки на основе стиля
            style = self._styles.get((row, col), {})
            font_color_hex = style.get('font_color')
            if font_color_hex:
                return QBrush(QColor(f"#{font_color_hex}"))
        elif role == Qt.ItemDataRole.FontRole:
            # TODO: Вернуть QFont для ячейки на основе стиля
            style = self._styles.get((row, col), {})
            font_attrs = {
                'bold': style.get('bold'),
                'italic': style.get('italic'),
                'underline': style.get('underline'),
                'font_size': style.get('font_size'),
                'font_name': style.get('font_name')
            }
            # Создаём QFont и применяем атрибуты
            font = QFont()
            if font_attrs['font_name']:
                font.setFamily(font_attrs['font_name'])
            if font_attrs['font_size']:
                font.setPointSize(int(font_attrs['font_size']))
            if font_attrs['bold'] is not None:
                font.setBold(bool(font_attrs['bold']))
            if font_attrs['italic'] is not None:
                font.setItalic(bool(font_attrs['italic']))
            # underline может быть строкой ('single', 'double', ...), обработка требует уточнения
            # if font_attrs['underline']:
            #     font.setUnderline(True) # Упрощённо
            return font
        elif role == Qt.ItemDataRole.TextAlignmentRole:
            # TODO: Вернуть флаги выравнивания на основе стиля
            style = self._styles.get((row, col), {})
            align_h = style.get('align', 'left') # 'left', 'center', 'right', 'fill', 'justify', 'distributed'
            align_v = style.get('valign', 'top') # 'top', 'vcenter', 'bottom', 'vjustify', 'vdistributed'
            # Сопоставление строк с флагами Qt
            align_map_h = {
                'left': Qt.AlignmentFlag.AlignLeft,
                'center': Qt.AlignmentFlag.AlignHCenter,
                'right': Qt.AlignmentFlag.AlignRight,
                'fill': Qt.AlignmentFlag.AlignJustify,
                'justify': Qt.AlignmentFlag.AlignJustify,
                'distributed': Qt.AlignmentFlag.AlignJustify,
            }
            align_map_v = {
                'top': Qt.AlignmentFlag.AlignTop,
                'vcenter': Qt.AlignmentFlag.AlignVCenter,
                'bottom': Qt.AlignmentFlag.AlignBottom,
                'vjustify': Qt.AlignmentFlag.AlignJustify,
                'vdistributed': Qt.AlignmentFlag.AlignJustify,
            }
            h_flag = align_map_h.get(align_h, Qt.AlignmentFlag.AlignLeft)
            v_flag = align_map_v.get(align_v, Qt.AlignmentFlag.AlignTop)
            return h_flag | v_flag

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        """Возвращает данные для заголовков строк или столбцов."""
        if role != Qt.ItemDataRole.DisplayRole:
            return None

        if orientation == Qt.Orientation.Horizontal:
            if 0 <= section < len(self._headers):
                return self._headers[section]
        elif orientation == Qt.Orientation.Vertical:
            if 0 <= section < len(self._row_headers):
                return self._row_headers[section]
        return None

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.ItemDataRole.EditRole) -> bool:
        """
        Устанавливает данные для указанной ячейки.
        Вызывается при редактировании в QTableView.
        """
        # TODO: Реализовать изменение данных через AppController
        # 1. Проверить, что index.isValid() и role == Qt.EditRole
        # 2. Преобразовать index.row(), index.column() в адрес ячейки Excel (e.g., 'A1')
        # 3. Вызвать метод AppController для обновления значения ячейки
        #    app_controller.update_cell_value(self.sheet_name, cell_address, value)
        # 4. Обновить внутреннее состояние модели (self._data)
        # 5. Вызвать self.dataChanged.emit() для уведомления представления
        #    self.dataChanged.emit(index, index, [role])
        # logger.debug(f"Попытка установки данных в ячейку ({index.row()}, {index.column()}) значение {value}")
        if not index.isValid() or role != Qt.ItemDataRole.EditRole:
            return False

        row = index.row()
        col = index.column()
        if row >= len(self._data) or col >= len(self._data[0]):
            return False

        cell_address = self._index_to_cell_address(row, col)
        try:
            # Вызов AppController для обновления ячейки
            success = self.app_controller.update_cell_value(self.sheet_name, cell_address, value)
            if success:
                # Обновляем локальное состояние модели
                self._data[row][col] = value
                # Уведомляем представление об изменении
                self.dataChanged.emit(index, index, [role])
                logger.info(f"Ячейка {cell_address} обновлена через AppController.")
                return True
            else:
                logger.error(f"AppController не смог обновить ячейку {cell_address}.")
                return False
        except Exception as e:
            logger.error(f"Ошибка при обновлении ячейки {cell_address} через AppController: {e}", exc_info=True)
            return False

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        """Определяет флаги для ячейки (редактируемая, доступная и т.д.)."""
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags

        # Делаем ячейки редактируемыми
        return super().flags(index) | Qt.ItemFlag.ItemIsEditable

    def _index_to_cell_address(self, row: int, col: int) -> str:
        """
        Преобразует индексы строки и столбца (0-based) в адрес ячейки Excel (e.g., 'A1').
        """
        col_name = self._index_to_column_name(col)
        row_num = row + 1
        return f"{col_name}{row_num}"

    def span(self, row: int, column: int) -> tuple[int, int]:
        """
        Возвращает размер объединённой области для ячейки (row, column).
        Используется QTableView для отображения объединённых ячеек.
        """
        # TODO: Реализовать логику для объединённых ячеек
        # Нужно проверить, принадлежит ли ячейка (row, column) к объединённой области в self._merged_cells
        # Если да, вернуть (row_span, col_span), где row_span = bottom_row - top_row + 1, col_span = right_col - left_col + 1
        # Если нет, вернуть (1, 1)
        # logger.debug(f"Запрос span для ({row}, {column})")
        for top_row, left_col, bottom_row, right_col in self._merged_cells:
            if top_row <= row <= bottom_row and left_col <= column <= right_col:
                row_span = bottom_row - top_row + 1
                col_span = right_col - left_col + 1
                # logger.debug(f"Ячейка ({row}, {column}) принадлежит объединению ({top_row}, {left_col})-({bottom_row}, {right_col}). Span: ({row_span}, {col_span})")
                return row_span, col_span
        # logger.debug(f"Ячейка ({row}, {column}) не объединена. Span: (1, 1)")
        return 1, 1
