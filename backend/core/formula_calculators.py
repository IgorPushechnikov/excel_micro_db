# backend/core/formula_calculators.py
"""
Модуль, содержащий функции для вычисления сложных формул,
аналогичных тем, что используются в Excel, но реализованных на Python.

Каждая функция получает данные через DataManager и сохраняет результат обратно.
"""

import pandas as pd
from datetime import datetime
import calendar
import logging
from typing import Dict, Any, List, Optional
from backend.utils.logger import get_logger

logger = get_logger(__name__)


def calculate_age_string(start_date_cell_value: Any, end_date_cell_value: Any) -> str:
    """
    Вычисляет строку возраста по формату Excel-формулы.

    Args:
        start_date_cell_value: Значение ячейки с начальной датой (ожидается datetime, строка в формате даты или pd.Timestamp).
        end_date_cell_value: Значение ячейки с конечной датой (ожидается datetime, строка в формате даты или pd.Timestamp).

    Returns:
        str: Строка с возрастом, например, "5 лет 2 месяца 10 дней".
             Возвращает пустую строку, если дата начала позже даты окончания или если даты недействительны.
    """
    try:
        # Приведение к datetime, если строка или pd.Timestamp
        start_date = _convert_to_datetime(start_date_cell_value)
        end_date = _convert_to_datetime(end_date_cell_value)

        if start_date is None or end_date is None:
            if logger.isEnabledFor(logging.WARNING):
                logger.warning(f"Неверный формат даты: start={start_date_cell_value}, end={end_date_cell_value}")
            return ""

        # Проверка, что начальная дата не позже конечной
        if start_date > end_date:
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"Начальная дата ({start_date}) позже конечной ({end_date})")
            return ""

        # Вычисляем годы, месяцы, дни вручную
        years = end_date.year - start_date.year
        months = end_date.month - start_date.month
        days = end_date.day - start_date.day

        # Корректировка дней и месяцев, если день в начальной дате больше, чем в конечной
        if days < 0:
            months -= 1
            # Получаем количество дней в предыдущем месяце конечной даты
            if end_date.month == 1:
                prev_month = 12
                prev_year = end_date.year - 1
            else:
                prev_month = end_date.month - 1
                prev_year = end_date.year
            days_in_prev_month = calendar.monthrange(prev_year, prev_month)[1]
            days += days_in_prev_month

        # Корректировка месяцев, если месяц в начальной дате больше, чем в конечной
        if months < 0:
            years -= 1
            months += 12

        # Функция для определения окончания
        def get_ending(value: int, unit_type: str) -> str:
            if value == 0:
                return ""
            last_digit = value % 10
            last_two_digits = value % 100

            if unit_type == "year":
                if 11 <= last_two_digits <= 14 or last_digit == 0 or 5 <= last_digit <= 9:
                    return " лет"
                elif 2 <= last_digit <= 4:
                    return " года"
                elif last_digit == 1:
                    return " год"
            elif unit_type == "month":
                if 11 <= last_two_digits <= 14 or last_digit == 0 or 5 <= last_digit <= 9:
                    return " месяцев"
                elif 2 <= last_digit <= 4:
                    return " месяца"
                elif last_digit == 1:
                    return " месяц"
            elif unit_type == "day":
                if 11 <= last_two_digits <= 14 or last_digit == 0 or 5 <= last_digit <= 9:
                    return " дней"
                elif 2 <= last_digit <= 4:
                    return " дня"
                elif last_digit == 1:
                    return " день"
            return ""

        year_str = f"{years}{get_ending(years, 'year')}" if years > 0 else ""
        month_str = f"{months}{get_ending(months, 'month')}" if months > 0 else ""
        day_str = f"{days}{get_ending(days, 'day')}" if days > 0 else ""

        result_parts = [s for s in [year_str, month_str, day_str] if s]
        result = " ".join(result_parts)
        return result.strip() # Убираем лишние пробелы в начале/конце

    except Exception as e:
        logger.error(f"Ошибка при вычислении возраста: {e}", exc_info=True)
        return "" # Возвращаем пустую строку в случае ошибки


def _convert_to_datetime(value: Any) -> Optional[datetime]:
    """
    Преобразует значение в datetime. Поддерживает datetime, строку, pd.Timestamp.

    Args:
        value: Значение для преобразования.

    Returns:
        datetime: Преобразованное значение или None, если преобразование не удалось.
    """
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, str):
        try:
            # Пытаемся преобразовать строку в дату
            dt = pd.to_datetime(value, errors='coerce')
            if pd.isna(dt):
                return None
            return dt.to_pydatetime()
        except Exception:
            return None
    # Если тип не поддерживается
    return None


def apply_age_formula_to_column(
    data_manager,
    sheet_name: str,
    start_date_cell_addr: str, # e.g., 'A3'
    end_date_cell_addr: str,   # e.g., 'C3'
    result_column_addr: str    # e.g., 'F' (will be applied to F3, F4, ...)
):
    """
    Применяет формулу возраста к столбцу.
    Берёт даты из start_date_cell_addr и end_date_cell_addr,
    вычисляет строку возраста и записывает её в result_column_addr,
    начиная с той же строки, что и даты, до конца заполненных данных.

    Args:
        data_manager: Экземпляр DataManager.
        sheet_name (str): Имя листа.
        start_date_cell_addr (str): Адрес ячейки с начальной датой (e.g., 'A3').
        end_date_cell_addr (str): Адрес ячейки с конечной датой (e.g., 'C3').
        result_column_addr (str): Буква результирующего столбца (e.g., 'F').

    Returns:
        bool: True, если успешно выполнено.
    """
    try:
        logger.info(f"Начало применения формулы возраста к листу '{sheet_name}', столбец {result_column_addr}, даты из {start_date_cell_addr} и {end_date_cell_addr}.")

        # Получаем сырые и редактируемые данные листа
        raw_data_rows, editable_data_rows = data_manager.get_sheet_data(sheet_name)
        
        # Создаем словарь для быстрого доступа к редактируемым значениям
        editable_dict = {item["cell_address"]: item["value"] for item in editable_data_rows}

        # Найдем значения дат в указанных ячейках
        start_date_val = editable_dict.get(start_date_cell_addr.upper())
        end_date_val = editable_dict.get(end_date_cell_addr.upper())

        if logger.isEnabledFor(logging.DEBUG):
            logger.debug(f"Найдены даты: {start_date_cell_addr}={start_date_val}, {end_date_cell_addr}={end_date_val}")

        # Проверим, найдены ли обе даты
        if start_date_val is None or end_date_val is None:
            logger.error(f"Не удалось найти значения дат в ячейках {start_date_cell_addr} или {end_date_cell_addr}.")
            return False

        # Извлекаем номер строки из адресов (Excel-нумерация с 1, Python с 0)
        def row_from_address(addr: str) -> int:
            row_part = "".join(filter(str.isdigit, addr))
            return int(row_part) - 1 if row_part.isdigit() else -1

        start_row_idx = row_from_address(start_date_cell_addr)
        end_row_idx = row_from_address(end_date_cell_addr)
        result_start_row_idx = min(start_row_idx, end_row_idx) if start_row_idx >= 0 and end_row_idx >= 0 else -1

        if result_start_row_idx < 0:
            logger.error(f"Неверный формат адреса строки в {start_date_cell_addr} или {end_date_cell_addr}.")
            return False

        # Определяем максимальную строку, участвующую в editable_data
        max_row_idx = -1
        for item in editable_data_rows:
            addr = item["cell_address"]
            row_idx = row_from_address(addr)
            if row_idx > max_row_idx:
                max_row_idx = row_idx

        # Если editable_data пуст или не содержит строк, можно выйти
        if max_row_idx < 0:
            logger.warning(f"Лист '{sheet_name}' не содержит редактируемых данных. Нечего обновлять.")
            return False

        # Проходим по строкам, начиная с result_start_row_idx, до max_row_idx
        success_count = 0
        for i in range(result_start_row_idx, max_row_idx + 1):
            # Для каждой строки, мы всё равно используем одни и те же start_date_val и end_date_val
            calculated_result = calculate_age_string(start_date_val, end_date_val)
            result_cell_addr = f"{result_column_addr.upper()}{i + 1}" # Обратно в Excel-нумерацию

            # Обновляем значение в БД
            update_success = data_manager.update_cell_value(sheet_name, result_cell_addr, calculated_result)
            if update_success:
                success_count += 1
                # Проверяем уровень перед логированием внутри цикла
                if logger.isEnabledFor(logging.DEBUG):
                    logger.debug(f"Обновлена ячейка {result_cell_addr} значением '{calculated_result}'")
            else:
                logger.error(f"Не удалось обновить ячейку {result_cell_addr}.")

        logger.info(f"Формула возраста применена. Обновлено {success_count} ячеек в столбце {result_column_addr.upper()} листа '{sheet_name}'.")
        return True

    except Exception as e:
        logger.error(f"Ошибка при применении формулы возраста к столбцу: {e}", exc_info=True)
        return False

# --- Пример использования (для тестирования) ---
# if __name__ == "__main__":
#     # Этот код будет выполняться только при прямом запуске этого файла
#     # и не будет выполняться при импорте.
#     import sys
#     sys.path.append("../../") # Позволяет импортировать из backend.core
#     from app_controller import create_app_controller
#     from storage.base import ProjectDBStorage
#     
#     # Пример данных
#     start_date_val = datetime(1990, 5, 15) # или строка "1990-05-15"
#     end_date_val = datetime(2023, 10, 7)
#     result = calculate_age_string(start_date_val, end_date_val)
#     print(f"Результат для {start_date_val} - {end_date_val}: '{result}'")

#     # Пример вызова apply_age_formula_to_column (требует реальный экземпляр DataManager)
#     # app_ctrl = create_app_controller("path_to_project")
#     # app_ctrl.initialize()
#     # if app_ctrl.storage:
#     #     success = apply_age_formula_to_column(
#     #         app_ctrl.data_manager,
#     #         "Sheet1",
#     #         "A3",
#     #         "C3",
#     #         "F"
#     #     )
#     #     print(f"Применение формулы: {'Успех' if success else 'Ошибка'}")
#     # else:
#     #     print("Не удалось подключиться к БД проекта для теста.")
