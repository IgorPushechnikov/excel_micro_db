# src/constructor/components/style_manager.py
"""
Модуль для управления стилями ячеек в новом GUI Excel Micro DB.
"""

import logging
from typing import Optional, Dict, Any, Tuple
from dataclasses import dataclass

from PySide6.QtGui import QColor, QFont

# Импортируем AppController
from src.core.app_controller import AppController

logger = logging.getLogger(__name__)


@dataclass
class CellStyle:
    """
    Класс для представления стиля ячейки.
    """
    bg_color: Optional[QColor] = None
    text_color: Optional[QColor] = None
    font: Optional[QFont] = None
    # alignment: Optional[Qt.AlignmentFlag] = None
    # border: Optional[QPen] = None


class StyleManager:
    """
    Менеджер стилей для SheetEditor.
    Отвечает за загрузку, кэширование и применение стилей к ячейкам.
    """

    def __init__(self, app_controller: AppController, sheet_name: str):
        """
        Инициализирует менеджер стилей.

        Args:
            app_controller (AppController): Экземпляр AppController.
            sheet_name (str): Имя листа.
        """
        self.app_controller = app_controller
        self.sheet_name = sheet_name
        self._style_cache: Dict[Tuple[int, int], CellStyle] = {}
        self._load_styles()

    def _load_styles(self):
        """
        Загружает стили для листа из AppController и кэширует их.
        """
        logger.info(f"Загрузка стилей для листа '{self.sheet_name}'...")
        try:
            # Предполагается, что AppController имеет метод get_sheet_styles
            # который возвращает словарь {(row, col): style_dict}
            # style_dict содержит ключи, соответствующие полям CellStyle
            styles_data = self.app_controller.get_sheet_styles(self.sheet_name)
            if styles_data:
                for (row, col), style_dict in styles_data.items():
                    # Конвертируем данные стиля из словаря в CellStyle
                    cell_style = self._dict_to_cell_style(style_dict)
                    self._style_cache[(row, col)] = cell_style
                logger.info(f"Загружено {len(self._style_cache)} стилей для листа '{self.sheet_name}'.")
            else:
                logger.debug(f"Стили для листа '{self.sheet_name}' не найдены.")
        except Exception as e:
            logger.error(f"Ошибка при загрузке стилей для листа '{self.sheet_name}': {e}", exc_info=True)

    def _dict_to_cell_style(self, style_dict: Dict[str, Any]) -> CellStyle:
        """
        Конвертирует словарь стиля в объект CellStyle.

        Args:
            style_dict (Dict[str, Any]): Словарь с данными стиля.

        Returns:
            CellStyle: Объект стиля.
        """
        cell_style = CellStyle()

        # Обработка фонового цвета
        bg_color_str = style_dict.get("bg_color")
        if bg_color_str:
            try:
                cell_style.bg_color = QColor(bg_color_str)
            except Exception as e:
                logger.warning(f"Неверный формат цвета фона '{bg_color_str}': {e}")

        # Обработка цвета текста
        text_color_str = style_dict.get("text_color")
        if text_color_str:
            try:
                cell_style.text_color = QColor(text_color_str)
            except Exception as e:
                logger.warning(f"Неверный формат цвета текста '{text_color_str}': {e}")

        # Обработка шрифта
        font_info = style_dict.get("font")
        if font_info:
            try:
                # Предполагаем, что font_info - это словарь с ключами 'family', 'point_size', 'bold', 'italic'
                font = QFont()
                font.setFamily(font_info.get("family", "Arial"))
                font.setPointSize(font_info.get("point_size", 10))
                font.setBold(font_info.get("bold", False))
                font.setItalic(font_info.get("italic", False))
                cell_style.font = font
            except Exception as e:
                logger.warning(f"Ошибка при создании шрифта из данных {font_info}: {e}")

        return cell_style

    def get_style(self, row: int, col: int) -> Optional[CellStyle]:
        """
        Возвращает стиль для ячейки (row, col).

        Args:
            row (int): Номер строки (0-базированный).
            col (int): Номер столбца (0-базированный).

        Returns:
            Optional[CellStyle]: Стиль ячейки или None, если стиль не задан.
        """
        return self._style_cache.get((row, col))

    def set_style(self, row: int, col: int, style: CellStyle) -> bool:
        """
        Устанавливает стиль для ячейки (row, col) и сохраняет его через AppController.

        Args:
            row (int): Номер строки (0-базированный).
            col (int): Номер столбца (0-базированный).
            style (CellStyle): Стиль для установки.

        Returns:
            bool: True, если стиль успешно установлен и сохранен, иначе False.
        """
        try:
            # Сохраняем стиль в кэш
            self._style_cache[(row, col)] = style

            # Конвертируем CellStyle обратно в словарь для сохранения
            style_dict = self._cell_style_to_dict(style)

            # Отправляем изменение в AppController
            # Предполагается, что AppController имеет метод update_sheet_cell_style
            success = self.app_controller.update_sheet_cell_style(self.sheet_name, row, col, style_dict)
            if success:
                logger.debug(f"Стиль для ячейки ({row}, {col}) успешно обновлен.")
                return True
            else:
                logger.error(f"Не удалось обновить стиль для ячейки ({row}, {col}) в проекте.")
                # Удаляем из кэша, если сохранение не удалось
                self._style_cache.pop((row, col), None)
                return False
        except Exception as e:
            logger.error(f"Ошибка при установке стиля для ячейки ({row}, {col}): {e}", exc_info=True)
            return False

    def _cell_style_to_dict(self, style: CellStyle) -> Dict[str, Any]:
        """
        Конвертирует объект CellStyle в словарь для сохранения.

        Args:
            style (CellStyle): Объект стиля.

        Returns:
            Dict[str, Any]: Словарь с данными стиля.
        """
        style_dict = {}

        if style.bg_color:
            style_dict["bg_color"] = style.bg_color.name()  # Сохраняем как hex-строку

        if style.text_color:
            style_dict["text_color"] = style.text_color.name()

        if style.font:
            style_dict["font"] = {
                "family": style.font.family(),
                "point_size": style.font.pointSize(),
                "bold": style.font.bold(),
                "italic": style.font.italic()
            }

        return style_dict