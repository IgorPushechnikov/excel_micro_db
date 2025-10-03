# backend/constructor/widgets/sheet_editor/cell_formatting.py
"""
Модуль для форматирования значений ячеек Excel на основе числового формата.
Используется в SheetDataModel для отображения значений в QTableView.
"""

import re
from datetime import datetime, date, time
from typing import Any, Optional
# ИСПРАВЛЕНО: Безопасный импорт функции проверки формата даты из openpyxl
try:
    # Попробуем импортировать функцию для проверки формата даты
    # В разных версиях openpyxl это может быть по-разному
    from openpyxl.styles.numbers import is_date_format as opx_is_date_format
except ImportError:
    # fallback если функция не найдена
    def opx_is_date_format(fmt_code):
        # Простая эвристика
        indicators = ['Y', 'M', 'D', 'H', 'S', 'AM', 'PM', 'A/P']
        if not fmt_code or not isinstance(fmt_code, str):
            return False
        fmt_upper = fmt_code.upper()
        return any(indicator in fmt_upper for indicator in indicators)

# Импортируем logger из utils
# ИСПРАВЛЕНО: Корректный путь к logger внутри backend
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
    Основано на openpyxl или эвристике.
    """
    # Проверяем, не является ли это одним из стандартных типов, связанных с датой
    # ИСПРАВЛЕНО: Используем правильную функцию из openpyxl или fallback
    try:
        if opx_is_date_format is not None and callable(opx_is_date_format):
            return opx_is_date_format(number_format_code)
        else:
            logger.debug("Функция opx_is_date_format не доступна, используем эвристику.")
    except Exception as e:
        logger.debug(f"Ошибка при вызове opx_is_date_format: {e}. Используем эвристику.")
    
    # fallback на эвристику, если openpyxl не помог
    if not number_format_code or not isinstance(number_format_code, str):
        return False
        
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
            # (логика упрощена в этом примере)
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
                # logger.debug(f"_apply_date_format: Formatting {value} with {python_format}")
                return value.strftime(python_format)
            else:
                logger.warning(f"Не удалось сопоставить формат даты '{number_format_code}' с Python strftime.")
                # Попробуем использовать стандартный формат
                return str(value)
        except (ValueError, TypeError) as e:
            logger.warning(f"Ошибка форматирования даты {value} с форматом {number_format_code}: {e}")
            return str(value) # Возврат строкового представления как fallback

    # --- НОВОЕ: Проверка, является ли значение строкой, похожей на дату ---
    # Если значение - строка, и формат помечен как дата, возможно, это ISO-дата
    if isinstance(value, str) and _is_date_format(number_format_code):
        # Попробуем распарсить строку как дату/время
        # Пример: "2024-02-20T00:00:00"
        try:
            # Попробуем несколько распространённых форматов
            possible_formats = [
                "%Y-%m-%dT%H:%M:%S",  # ISO 8601
                "%Y-%m-%d %H:%M:%S",  # Common datetime
                "%Y-%m-%d",           # Date only
                "%d.%m.%Y",           # Russian date format
                "%m/%d/%Y",           # US date format
                "%d/%m/%Y",           # UK date format
            ]
            parsed_dt = None
            for fmt in possible_formats:
                try:
                    parsed_dt = datetime.strptime(value, fmt)
                    break  # Если успешно, прекращаем перебор
                except ValueError:
                    continue  # Пробуем следующий формат
            
            if parsed_dt:
                # Если успешно распарсили, форматируем как дату
                python_format = _excel_date_format_to_python_strftime(number_format_code)
                if python_format:
                    # logger.debug(f"_apply_date_format: Formatting parsed datetime {parsed_dt} with {python_format}")
                    return parsed_dt.strftime(python_format)
                else:
                    logger.warning(f"Не удалось сопоставить формат даты '{number_format_code}' с Python strftime для строки '{value}'.")
                    # Попробуем использовать стандартный формат
                    return str(parsed_dt.date()) # Возвращаем только дату
            else:
                logger.debug(f"Не удалось распарсить строку '{value}' как дату/время для формата '{number_format_code}'.")
        except Exception as e:
            logger.warning(f"Ошибка при попытке распарсить строку '{value}' как дату: {e}")

    # Если значение не является датой и не строкой-датой, но формат помечен как дата, 
    # возможно, это число (день с 1900-01-01) - сложнее и требует знания о типе данных.
    # OpenPyxl обычно сам конвертирует.
    # Если значение - строка, и формат - дата, это странно.
    logger.debug(f"Значение '{value}' типа {type(value)} не является датой или строкой-датой, но формат '{number_format_code}' помечен как дата.")
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

    # --- НОВОЕ: Замена специфичных форматов на русские ---
    # Заменяем %Y-%m-%d на %d.%m.%y для русского формата
    # Это упрощённое правило. В реальности может потребоваться более сложная логика.
    # Например, если формат Excel был "DD.MM.YYYY", то python_format будет "%d.%m.%Y".
    # Но если формат Excel был "YYYY-MM-DD", то python_format будет "%Y-%m-%d".
    # Мы хотим привести к виду "03.10.25".
    # Проверим, содержит ли формат дату
    if '%Y' in python_format or '%y' in python_format:
        # Заменяем %Y на %y, если нужно двухзначный год
        # python_format = python_format.replace('%Y', '%y')
        # Заменяем разделители на точки
        python_format = python_format.replace('-', '.')
        python_format = python_format.replace('/', '.')
        # Заменяем порядок на DD.MM.YYYY или DD.MM.YY
        # Это сложно сделать универсально. Проще использовать конкретный формат.
        # Например, если формат содержит год, месяц и день, то форматируем как "DD.MM.YY"
        # Но это может быть не всегда корректно.
        # Попробуем более простой подход: заменить %Y-%m-%d на %d.%m.%y
        python_format = re.sub(r'%Y[-/.]%m[-/.]%d', '%d.%m.%y', python_format)
        python_format = re.sub(r'%d[-/.]%m[-/.]%Y', '%d.%m.%y', python_format)
        python_format = re.sub(r'%d[-/.]%m[-/.]%y', '%d.%m.%y', python_format)
        # Также заменить %y
        python_format = re.sub(r'%Y[-/.]%m[-/.]%y', '%d.%m.%y', python_format)
        
    # --- КОНЕЦ НОВОГО ---

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
    # ИСПРАВЛЕНО: Используем исправленную функцию проверки
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