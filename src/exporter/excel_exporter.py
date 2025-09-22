# src/exporter/excel_exporter.py
"""
Модуль-дирижёр для экспорта данных проекта Excel Micro DB в новый Excel-файл.
Координирует работу вспомогательных модулей экспорта: data_and_formulas, styles, charts.
Использует openpyxl для создания файла.
"""

import logging
from typing import Dict, Any, List, Optional
from pathlib import Path
import sys

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

logger = get_logger(__name__)

# Импортируем вспомогательные модули экспорта (предполагается, что они реализованы)
# ВАЖНО: Эти модули должны использовать openpyxl
try:
    from src.exporter import data_and_formulas_exporter
    from src.exporter import style_exporter
    from src.exporter import chart_exporter
    logger.debug("Вспомогательные модули экспорта (openpyxl) успешно импортированы.")
except ImportError as e:
    logger.critical(f"Не удалось импортировать вспомогательные модули экспорта: {e}")
    # Можно выбросить исключение или продолжить с ограниченной функциональностью
    # raise
    data_and_formulas_exporter = None
    style_exporter = None
    chart_exporter = None

# Импортируем openpyxl
try:
    from openpyxl import Workbook
    from openpyxl.workbook.workbook import Workbook as OpenpyxlWorkbook
    from openpyxl.worksheet.worksheet import Worksheet
    logger.debug("openpyxl успешно импортирован.")
except ImportError as e:
    logger.critical(f"Не удалось импортировать openpyxl: {e}")
    raise


def export_project_from_db(db_path: str, output_path: str) -> bool:
    """
    Экспортирует проект из SQLite БД в файл Excel (.xlsx).
    Загружает все данные проекта через storage.get_all_data() и делегирует
    экспорт вспомогательным модулям, используя openpyxl.

    Args:
        db_path (str): Путь к файлу БД проекта (.sqlite).
        output_path (str): Путь к файлу Excel, который будет создан.

    Returns:
        bool: True, если экспорт прошёл успешно, иначе False.
    """
    logger.info("=== НАЧАЛО ЭКСПОРТА ПРОЕКТА ИЗ БД (openpyxl - Дирижёр) ===")
    logger.info(f"Путь к БД проекта: {db_path}")
    logger.info(f"Путь к выходному файлу: {output_path}")

    db_path_obj = Path(db_path)
    output_path_obj = Path(output_path)

    if not db_path_obj.exists():
        logger.error(f"Файл БД проекта не найден: {db_path}")
        return False

    workbook: Optional[OpenpyxlWorkbook] = None
    try:
        # Импорт внутри блока try, чтобы избежать проблем при импорте модуля
        from src.storage.base import ProjectDBStorage

        # --- Загрузка всех данных проекта ---
        logger.debug("Подключение к БД проекта и загрузка всех данных...")
        with ProjectDBStorage(str(db_path_obj)) as storage:
            logger.debug("Подключение к БД проекта установлено через ProjectDBStorage.")
            project_data = storage.get_all_data()
            if not project_data:
                 logger.error("Не удалось загрузить данные проекта из БД.")
                 return False
            logger.info("Данные проекта успешно загружены из БД.")
        # ----------------------------------

        # --- Создание новой книги Excel ---
        workbook = Workbook()
        logger.info("Создана новая книга Excel (openpyxl).")

        # Удаляем дефолтный лист, если он есть
        if "Sheet" in workbook.sheetnames:
            workbook.remove(workbook["Sheet"])
            logger.debug("Удален дефолтный лист 'Sheet'.")

        sheets_data = project_data.get("sheets", {})

        if not sheets_data:
             logger.warning("В проекте не найдено листов. Создается пустой файл.")
             workbook.create_sheet("EmptySheet")
             workbook.save(str(output_path_obj))
             logger.info(f"Пустой файл сохранен: {output_path}")
             logger.info("=== ЭКСПОРТ ПРОЕКТА ИЗ БД (openpyxl - Дирижёр) ЗАВЕРШЕН ===")
             return True

        # --- Экспорт каждого листа ---
        logger.info(f"Найдено {len(sheets_data)} листов для экспорта.")
        for sheet_name, sheet_data in sheets_data.items():
            logger.info(f"Экспорт листа: {sheet_name}")
            
            # 1. Создаем лист в новой книге
            worksheet = workbook.create_sheet(title=sheet_name)
            logger.debug(f"Создан лист '{sheet_name}' в книге.")

            # 2. Делегируем экспорт данных и формул
            # Предполагаем, что data_and_formulas_exporter.export_sheet_data_and_formulas
            # принимает (workbook, worksheet, sheet_data)
            if data_and_formulas_exporter:
                try:
                    logger.debug(f"Делегирование экспорта данных/формул для листа '{sheet_name}'...")
                    data_and_formulas_exporter.export_sheet_data_and_formulas(workbook, worksheet, sheet_data)
                    logger.debug(f"Экспорт данных/формул для листа '{sheet_name}' завершен.")
                except Exception as e:
                    logger.error(f"Ошибка при экспорте данных/формул для листа '{sheet_name}': {e}", exc_info=True)
                    # Можно продолжить с другими листами или прервать
                    # continue
                    # raise
            else:
                logger.warning("Модуль data_and_formulas_exporter не доступен. Экспорт данных/формул пропущен.")

            # 3. Делегируем экспорт стилей
            # Предполагаем, что style_exporter.export_sheet_styles
            # принимает (workbook, worksheet, sheet_data)
            if style_exporter:
                try:
                    logger.debug(f"Делегирование экспорта стилей для листа '{sheet_name}'...")
                    style_exporter.export_sheet_styles(workbook, worksheet, sheet_data)
                    logger.debug(f"Экспорт стилей для листа '{sheet_name}' завершен.")
                except Exception as e:
                    logger.error(f"Ошибка при экспорте стилей для листа '{sheet_name}': {e}", exc_info=True)
                    # continue
            else:
                logger.warning("Модуль style_exporter не доступен. Экспорт стилей пропущен.")

            # 4. Делегируем экспорт диаграмм (если реализовано)
            # Предполагаем, что chart_exporter.export_sheet_charts
            # принимает (workbook, worksheet, sheet_data)
            if chart_exporter:
                try:
                    logger.debug(f"Делегирование экспорта диаграмм для листа '{sheet_name}'...")
                    chart_exporter.export_sheet_charts(workbook, worksheet, sheet_data)
                    logger.debug(f"Экспорт диаграмм для листа '{sheet_name}' завершен.")
                except AttributeError:
                    logger.warning("Функция export_sheet_charts не найдена в chart_exporter. Экспорт диаграмм пропущен.")
                except Exception as e:
                    logger.error(f"Ошибка при экспорте диаграмм для листа '{sheet_name}': {e}", exc_info=True)
            else:
                logger.warning("Модуль chart_exporter не доступен. Экспорт диаграмм пропущен.")

        # --- Сохранение файла ---
        logger.debug(f"Сохранение файла Excel в {output_path}...")
        workbook.save(str(output_path_obj))
        workbook.close()
        logger.info(f"Файл Excel успешно сохранен: {output_path}")
        logger.info("=== ЭКСПОРТ ПРОЕКТА ИЗ БД (openpyxl - Дирижёр) ЗАВЕРШЕН ===")
        return True

    except Exception as e:
        logger.error(f"Ошибка при экспорте проекта в файл '{output_path}': {e}", exc_info=True)
        # Пытаемся закрыть workbook, если он был создан
        try:
            if workbook is not None:
                workbook.close()
        except Exception as close_error:
             logger.error(f"Ошибка при закрытии книги Excel: {close_error}")
        return False

# Точка входа для тестирования напрямую
if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description="Экспорт проекта Excel Micro DB напрямую из БД с использованием openpyxl (Дирижёр).",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("db_path", help="Путь к файлу project_data.db")
    parser.add_argument("output_path", help="Путь для сохранения выходного .xlsx файла")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
                        help="Уровень логгирования (по умолчанию INFO)")

    args = parser.parse_args()

    # Настройка логирования для прямого запуска
    log_level = getattr(logging, args.log_level.upper())
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_formatter = logging.Formatter(log_format)
    logger.handlers.clear()
    logger.addHandler(console_handler)
    logger.setLevel(log_level)

    logger.info("=== ЗАПУСК СКРИПТА ЭКСПОРТА (openpyxl - Дирижёр) ===")

    success = export_project_from_db(args.db_path, args.output_path)

    if success:
        logger.info(f"Экспорт успешно завершен. Файл сохранен в: {args.output_path}")
        sys.exit(0)
    else:
        logger.error(f"Экспорт завершился с ошибкой.")
        sys.exit(1)
