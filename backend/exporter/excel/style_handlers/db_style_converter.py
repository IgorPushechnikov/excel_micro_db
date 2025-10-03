# backend/exporter/excel/style_handlers/db_style_converter.py
"""
Модуль для конвертации JSON-описаний стилей из БД в формат, понятный xlsxwriter.
"""

import json
import logging
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)


def json_style_to_xlsxwriter_format(style_json_str: str) -> Optional[Dict[str, Any]]:
    """
    Конвертирует JSON-строку стиля в словарь атрибутов xlsxwriter.

    Args:
        style_json_str (str): JSON-строка, содержащая атрибуты стиля.

    Returns:
        Optional[Dict[str, Any]]: Словарь с атрибутами формата xlsxwriter или None, если стиль пуст или неверен.
    """
    if not style_json_str:
        return None

    try:
        style_dict = json.loads(style_json_str)
    except json.JSONDecodeError as e:
        logger.error(f"Ошибка разбора JSON стиля: {e}")
        return None

    xlsxwriter_format = {}

    # --- Шрифт ---
    if 'font' in style_dict:
        font_data = style_dict['font']
        # xlsxwriter font properties: font_name, font_size, bold, italic, underline, font_strikeout, font_script, font_outline, font_shadow, font_color
        if 'name' in font_data:
            xlsxwriter_format['font_name'] = font_data['name']
        if 'sz' in font_data: # Размер шрифта
            xlsxwriter_format['font_size'] = float(font_data['sz'])
        if 'b' in font_data: # Жирный
            xlsxwriter_format['bold'] = bool(font_data['b'])
        if 'i' in font_data: # Курсив
            xlsxwriter_format['italic'] = bool(font_data['i'])
        # underline: 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting'
        if 'u' in font_data:
            xlsxwriter_format['underline'] = font_data['u']
        # strikeout
        if 'strike' in font_data:
            xlsxwriter_format['font_strikeout'] = bool(font_data['strike'])
        # color
        if 'color' in font_data and 'rgb' in font_data['color']:
            # xlsxwriter ожидает RGB в формате 'FF0000' (без '#')
            rgb_val = font_data['color']['rgb']
            if isinstance(rgb_val, str) and len(rgb_val) == 6:
                xlsxwriter_format['font_color'] = rgb_val
            elif isinstance(rgb_val, str) and len(rgb_val) == 8:
                # Убираем первый символ, если это alpha-канал (например, 'FF000000')
                xlsxwriter_format['font_color'] = rgb_val[2:]

    # --- Заливка (Pattern Fill) ---
    if 'fill' in style_dict:
        fill_data = style_dict['fill']
        # xlsxwriter fill properties: pattern, bg_color, fg_color
        # patternType: 'none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'
        if 'patternType' in fill_data:
            pt = fill_data['patternType']
            if pt == 'solid':
                # Для solid, проверим, есть ли цвет. Если нет, не устанавливаем pattern.
                if ('fgColor' in fill_data and 'rgb' in fill_data['fgColor']) or ('bgColor' in fill_data and 'rgb' in fill_data['bgColor']):
                    xlsxwriter_format['pattern'] = 1
            elif pt == 'none':
                xlsxwriter_format['pattern'] = 0
            elif pt is None:
                # Если patternType = null, не устанавливаем pattern, даже если цвета указаны (они могут быть прозрачными по умолчанию).
                pass
            else:
                # Сопоставление других типов паттернов (не все поддерживаются один-к-одному)
                # Устанавливаем solid (1) только если есть цвет (fg или bg), иначе оставляем без заливки (0 или не устанавливаем).
                if ('fgColor' in fill_data and 'rgb' in fill_data['fgColor']) or ('bgColor' in fill_data and 'rgb' in fill_data['bgColor']):
                     xlsxwriter_format['pattern'] = 1
                # else: # Не устанавливаем pattern, если нет цвета для неподдерживаемого типа паттерна

        # Цвета заливки
        if 'fgColor' in fill_data and 'rgb' in fill_data['fgColor']:
            rgb_val = fill_data['fgColor']['rgb']
            if isinstance(rgb_val, str) and len(rgb_val) == 6:
                xlsxwriter_format['fg_color'] = rgb_val
            elif isinstance(rgb_val, str) and len(rgb_val) == 8:
                xlsxwriter_format['fg_color'] = rgb_val[2:]
        if 'bgColor' in fill_data and 'rgb' in fill_data['bgColor']:
            rgb_val = fill_data['bgColor']['rgb']
            if isinstance(rgb_val, str) and len(rgb_val) == 6:
                xlsxwriter_format['bg_color'] = rgb_val
            elif isinstance(rgb_val, str) and len(rgb_val) == 8:
                xlsxwriter_format['bg_color'] = rgb_val[2:]

    # --- Границы ---
    if 'border' in style_dict:
        border_data = style_dict['border']
        # xlsxwriter border properties: border, border_color, bottom, bottom_color, top, top_color, left, left_color, right, right_color
        # Стили границ: 'none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'
        for side_name in ['left', 'right', 'top', 'bottom']:
            if side_name in border_data:
                side_data = border_data[side_name]
                # Стиль границы
                if 'style' in side_data:
                    style_val = side_data['style']
                    # Сопоставление стилей (упрощённое)
                    if style_val in ['thin', 'hair']:
                        xlsxwriter_format[f'{side_name}'] = 1
                    elif style_val == 'medium':
                        xlsxwriter_format[f'{side_name}'] = 2
                    elif style_val == 'thick':
                        xlsxwriter_format[f'{side_name}'] = 3
                    elif style_val == 'dashed':
                        xlsxwriter_format[f'{side_name}'] = 5
                    elif style_val == 'dotted':
                        xlsxwriter_format[f'{side_name}'] = 7
                    else:
                        # 'none' или другие -> 0
                        xlsxwriter_format[f'{side_name}'] = 0
                # Цвет границы
                if 'color' in side_data and 'rgb' in side_data['color']:
                    rgb_val = side_data['color']['rgb']
                    if isinstance(rgb_val, str) and len(rgb_val) == 6:
                        xlsxwriter_format[f'{side_name}_color'] = rgb_val
                    elif isinstance(rgb_val, str) and len(rgb_val) == 8:
                        xlsxwriter_format[f'{side_name}_color'] = rgb_val[2:]

    # --- Выравнивание ---
    if 'alignment' in style_dict:
        alignment_data = style_dict['alignment']
        # xlsxwriter alignment properties: align, valign, text_wrap, rotation, indent, shrink
        if 'horizontal' in alignment_data:
            h_val = alignment_data['horizontal']
            # Сопоставление ('left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed')
            if h_val in ['left', 'right', 'center', 'justify', 'distributed']:
                xlsxwriter_format['align'] = h_val
            elif h_val == 'fill':
                xlsxwriter_format['align'] = 'fill'
        if 'vertical' in alignment_data:
            v_val = alignment_data['vertical']
            # Сопоставление ('top', 'vcenter', 'bottom', 'vjustify', 'vdistributed')
            if v_val in ['top', 'vcenter', 'bottom', 'justify', 'distributed']:
                xlsxwriter_format['valign'] = v_val.replace('v', '') # xlsxwriter использует 'top', 'vcenter' -> 'vcenter', 'bottom', 'justify', 'distributed'
                if v_val == 'vcenter':
                    xlsxwriter_format['valign'] = 'vcenter'
                elif v_val == 'vjustify':
                    xlsxwriter_format['valign'] = 'vjustify'
                elif v_val == 'vdistributed':
                    xlsxwriter_format['valign'] = 'vdistributed'
                else:
                    # Для 'top', 'bottom' оставляем как есть
                    pass
        if 'wrapText' in alignment_data:
            xlsxwriter_format['text_wrap'] = bool(alignment_data['wrapText'])
        if 'textRotation' in alignment_data:
            # xlsxwriter принимает значения от -90 до 90, или 270 для вертикального текста
            rotation_val = int(alignment_data['textRotation'])
            if -90 <= rotation_val <= 90 or rotation_val == 270:
                xlsxwriter_format['rotation'] = rotation_val
        if 'indent' in alignment_data:
            xlsxwriter_format['indent'] = int(alignment_data['indent'])
        if 'shrinkToFit' in alignment_data:
            xlsxwriter_format['shrink'] = bool(alignment_data['shrinkToFit'])

    # --- Числовой формат ---
    if 'number_format' in style_dict:
        # --- ИЗМЕНЕНИЕ: Заменяем формат даты ---
        original_format = style_dict['number_format']
        # Заменяем mm-dd-yy на dd.mm.yyyy
        if original_format == 'mm-dd-yy':
            xlsxwriter_format['num_format'] = 'dd.mm.yyyy'
            logger.debug(f"Заменён формат даты: '{original_format}' -> 'dd.mm.yyyy'")
        else:
            xlsxwriter_format['num_format'] = original_format
        # --- КОНЕЦ ИЗМЕНЕНИЯ ---

    # --- Защита (обычно не применяется через формат ячейки в xlsxwriter напрямую) ---
    # xlsxwriter управляет защитой листа, а не отдельных ячеек, через worksheet.protect_range() и атрибуты листа.
    # if 'protection' in style_dict:
    #    prot_data = style_dict['protection']
    #    # locked, hidden - влияют на поведение при защите листа
    #    # Это сложнее и, возможно, требует отдельной логики на уровне листа.

    logger.debug(f"Конвертирован стиль: {style_dict} -> {xlsxwriter_format}")
    return xlsxwriter_format if xlsxwriter_format else None
