# backend/constructor/widgets/cell_formatting.py
"""
Модуль для форматирования значений ячеек Excel на основе числового формата.
Используется в SheetDataModel для отображения значений в QTableView.
"""

import re
from datetime import datetime, date, time
from typing import Any, Optional
import openpyxl.styles.numbers as openpyxl_number_formats
# Импортируем logger из utils
from utils.logger import get_logger # <-- ИСПРАВЛЕНО: было from src.utils.logger

logger = get_logger(__name__)

# Словарь для сопоставления общих кодов формата Excel с Python strftime
# Эти коды часто используются в пользовательских форматах
EXCEL_TO_PYTHON_DATE_FORMAT_MAP = {
    'DD': '%d', 'D': '%-d',  # День месяца (01-31, 1-31)
    'MM': '%m', 'M': '%-m',  # Месяц (01-12, 1-12)
    'YYYY': '%Y', 'YY': '%y', # Год (4 цифры, 2 цифры)
    'HH': '%H', 'H': '%-H',  # Час (00-23, 0-23)
    'MM_minute': '%M', 'M_minute': '%-M', # Минуты (00-59, 0-59) - требует контекста
    'SS': '%S', 'S': '%-S',  # Секунды (00-59, 0-59)
    'AM/PM': '%p', # AM/PM
    'A/P': '%p',   # A/P (заменяется на AM/PM)
}

def _is_date_format(number_format_code: str) -> bool:
    """
    Проверяет, является ли код числового формата форматом даты/времени.
    Основано на эвристике и часто используемых символах.
    """
    # Проверяем, не является ли это одним из стандартных типов, связанных с датой
    # openpyxl.classify_number_format возвращает 'Date', 'Time', 'DateTime' и т.д.
    format_type = openpyxl_number_formats.classify_number_format(number_format_code)
    if format_type in ['Date', 'Time', 'DateTime']:
        return True

    # Проверяем на наличие общих признаков даты/времени в пользовательских форматах
    # Используем версию с MM_minute для минут, чтобы избежать конфликта с месяцем
    date_time_indicators = ['Y', 'M', 'D', 'H', 'S', 'AM', 'PM', 'A/P']
    # Создаем строку для проверки, заменяя MM_minute на M_minute на время проверки
    check_str = number_format_code.upper()
    # Заменяем MM_minute и M_minute на метки, чтобы не мешали
    check_str = re.sub(r'MM(?![\w])', 'MONTH', check_str) # MM как месяц
    check_str = re.sub(r'M(?![\w])', 'MINUTE_OR_MONTH', check_str) # M как минута или месяц

    for indicator in date_time_indicators:
        if indicator.upper() in check_str:
            # Если встречается 'M', нужно проверить контекст, чтобы отличить от минут
            if indicator == 'M':
                # Проверим, есть ли рядом H или S, что указывает на минуты
                # Это упрощённая проверка
                m_pos = check_str.find('MINUTE_OR_MONTH')
                if m_pos != -1:
                    # Проверим окрестности M на наличие H или S
                    context = check_str[max(0, m_pos-5):m_pos+5]
                    if 'H' in context or 'S' in context:
                        # Возможно, это минуты, но если M встречается и как месяц, всё равно считаем датой
                        # Лучше уж показать как дату, чем как число, если не уверен
                        pass # Продолжаем проверку других индикаторов
            return True
    return False

def _apply_date_format(value: Any, number_format_code: str) -> Optional[str]:
    """
    Применяет формат даты к значению.
    Возвращает отформатированную строку или None, если форматирование невозможно.
    """
    if value is None:
        return None

    # Проверяем, является ли значение объектом даты/времени
    if isinstance(value, (datetime, date, time)):
        # Если это datetime и формат не содержит времени, отображаем только дату
        if isinstance(value, datetime) and not _contains_time_format(number_format_code):
            # Если формат содержит только дату, используем только дату
            value = value.date()
        # Если это date и формат содержит время, добавляем 00:00:00?
        # Обычно date остается date, если формат не предполагает время.
        # Попробуем использовать openpyxl для форматирования
        try:
            # openpyxl.utils.format_number может помочь, но он для чисел
            # Для дат/времени лучше использовать strftime
            # Однако, openpyxl.number_to_string может быть полезен
            # Но он требует workbook. Попробуем вручную.
            # Сопоставим код формата Excel с Python strftime
            python_format = _excel_date_format_to_python_strftime(number_format_code)
            if python_format:
                return value.strftime(python_format)
            else:
                logger.warning(f"Не удалось сопоставить формат даты '{number_format_code}' с Python strftime.")
                # Попробуем использовать стандартный формат
                return str(value)
        except (ValueError, TypeError) as e:
            logger.warning(f"Ошибка форматирования даты {value} с форматом {number_format_code}: {e}")
            return str(value) # Возврат строкового представления как fallback

    # Если значение не является датой, но формат помечен как дата, возможно, это число (день с 1900-01-01)
    # Это сложнее и требует знания о типе данных. OpenPyxl обычно сам конвертирует.
    # Если значение - строка, и формат - дата, это странно.
    logger.debug(f"Значение '{value}' типа {type(value)} не является датой, но формат '{number_format_code}' помечен как дата.")
    return str(value)

def _contains_time_format(number_format_code: str) -> bool:
    """Проверяет, содержит ли формат время (H, M(минуты), S)."""
    # Заменяем MM на MONTH, чтобы не мешало минутам
    check_str = re.sub(r'MM(?![\w])', 'MONTH', number_format_code.upper())
    # Проверяем H и S
    if 'H' in check_str or 'S' in check_str:
        return True
    # Проверяем M (минуты), но не M (месяц)
    # Ищем M, который не является частью MM или MONTH
    # Регулярное выражение для M, окруженного не-буквами или началом/конец строки, и не после/до M
    # Это сложно. Проще проверить на общие комбинации.
    # Проверим, есть ли 'M' отдельно от 'MM' и 'MONTH'
    # Удалим MONTH и MM, и посмотрим, осталась ли M
    temp_str = re.sub(r'MM(?![\w])', '', check_str)
    temp_str = re.sub(r'MONTH(?![\w])', '', temp_str)
    # Теперь ищем M как минуту
    # Это не идеально, но лучше, чем ничего
    if re.search(r'(?<!M)M(?!M)', temp_str): # M, не предшествуемая и не за которой следует M
        return True
    return False

def _excel_date_format_to_python_strftime(excel_format: str) -> Optional[str]:
    """
    Пробует преобразовать код формата даты Excel в Python strftime строку.
    Это упрощённая версия, поддерживающая основные коды.
    """
    if not excel_format:
        return None

    # Проверим, не является ли это одним из стандартных форматов
    # openpyxl может помочь, но мы будем использовать эвристику
    # Сначала заменим MM (месяц) и HH, SS (время) на временные токены
    temp_format = excel_format.replace('MM', 'TEMP_MONTH_TOKEN')
    temp_format = temp_format.replace('HH', 'TEMP_HOUR_TOKEN')
    temp_format = temp_format.replace('SS', 'TEMP_SEC_TOKEN')

    # Теперь заменим M (минуты) на M_minute, если он окружен H или S
    # Это сложная логика. Проще заменить M_minute -> %M, M -> %m, но с осторожностью.
    # Проверим, есть ли H или S рядом с M
    # Регулярное выражение для M, окруженного H или S
    temp_format = re.sub(r'(?<=[HS])M(?=[HS])', 'TEMP_MINUTE_TOKEN', temp_format)
    temp_format = re.sub(r'(?<=[HS])M(?![A-Z])', 'TEMP_MINUTE_TOKEN', temp_format) # M после H/S
    temp_format = re.sub(r'(?<![A-Z])M(?=[HS])', 'TEMP_MINUTE_TOKEN', temp_format) # M до H/S

    # Теперь заменим оставшиеся M (месяцы) на %m
    python_format = temp_format.replace('M', '%m')
    # Заменим минуты
    python_format = python_format.replace('TEMP_MINUTE_TOKEN', '%M')
    # Заменим месяцы
    python_format = python_format.replace('TEMP_MONTH_TOKEN', '%m')
    # Заменим часы и секунды
    python_format = python_format.replace('TEMP_HOUR_TOKEN', '%H')
    python_format = python_format.replace('TEMP_SEC_TOKEN', '%S')

    # Теперь заменим DD, YYYY, YY, AM/PM
    python_format = python_format.replace('DD', '%d')
    python_format = python_format.replace('D', '%-d') # День без ведущего нуля
    python_format = python_format.replace('YYYY', '%Y')
    python_format = python_format.replace('YY', '%y')
    python_format = python_format.replace('AM/PM', '%p')
    python_format = python_format.replace('A/P', '%p')

    # Убираем кавычки и другие специальные символы, которые не влияют на strftime
    # Например, [RED], $, "текст"
    # Это требует сложной логики. Пока уберем простые кавычки и символы, не являющиеся форматами
    # Пример: "DD/MM/YYYY" -> %d/%m/%Y
    # Пример: #,##0.00" руб." -> не дата
    # Нам нужно убедиться, что это формат даты
    # Убираем двойные кавычки и их содержимое
    python_format = re.sub(r'\"[^\"]*\"', '', python_format)
    # Убираем квадратные скобки и их содержимое (например, [RED])
    python_format = re.sub(r'\[[^\]]*\]', '', python_format)

    # Убираем символы, не являющиеся спецификаторами формата или разделителями
    # Оставляем только % и символы, которые могут быть в strftime
    # Это может быть неточно. Проверим, содержит ли строка хотя бы один %.
    if '%' not in python_format:
         logger.warning(f"Преобразованный формат '{python_format}' не содержит спецификаторов. Исходный: '{excel_format}'")
         return None

    return python_format


def format_cell_value(value: Any, number_format_code: Optional[str]) -> str:
    """
    Форматирует значение ячейки на основе числового формата.

    Args:
        value: Значение ячейки (может быть str, int, float, datetime, date, time, None).
        number_format_code: Код числового формата Excel (например, 'DD.MM.YYYY', '0.00%', '#,##0').

    Returns:
        str: Отформатированное строковое представление значения.
             Возвращает str(value), если форматирование невозможно или не требуется.
    """
    if number_format_code is None:
        return str(value) if value is not None else ""

    # Проверяем, является ли формат форматом даты/времени
    if _is_date_format(number_format_code):
        formatted_date_str = _apply_date_format(value, number_format_code)
        if formatted_date_str is not None:
            return formatted_date_str
        else:
            # Если форматирование даты не удалось, возвращаем строку как есть
            logger.warning(f"Форматирование даты не удалось для значения {value} и формата '{number_format_code}'. Возвращаем строку.")

    # Здесь можно добавить логику для других форматов (числа, проценты и т.д.)
    # Пока что, если это не дата, и не реализовано, возвращаем строку
    # Для чисел можно использовать openpyxl.utils.format_number
    # Но это требует больше кода. Пока оставим как есть, если не дата.
    # Попробуем использовать openpyxl.number_to_string как fallback для чисел
    # Это требует workbook, которого у нас нет. Используем стандартный str.
    # Или можно попробовать использовать locale-aware форматирование, но это сложно.

    return str(value)
