# backend/importer/xlwings_importer.py
"""
Модуль для импорта данных из Excel-файла с помощью xlwings.
Поддерживает импорт:
- Сырых данных (значения, формулы)
- Стилей
- Диаграмм
- Метаданных листов
- Именованных диапазонов
и других элементов, доступных через xlwings.

Использует ProjectDBStorage для сохранения данных в БД.
"""
import logging
import os
from typing import Dict, Any, List, Optional, Callable # <-- Добавлен Callable
import json
import xlwings as xw

from backend.storage.base import ProjectDBStorage
from backend.utils.logger import get_logger

logger = get_logger(__name__)

def import_all_from_excel_xlwings(
    storage: ProjectDBStorage,
    file_path: str,
    progress_callback: Optional[Callable[[int, str], None]] = None # <-- НОВОЕ
) -> bool:
    """
    Импортирует все поддерживаемые данные из Excel-файла через xlwings в БД.

    Args:
        storage (ProjectDBStorage): Экземпляр хранилища БД.
        file_path (str): Путь к Excel-файлу (.xlsx, .xls).
        progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
            Принимает (процент: int, сообщение: str).

    Returns:
        bool: True, если импорт прошёл успешно, иначе False.
    """
    if not storage:
        logger.error("Экземпляр ProjectDBStorage не предоставлен.")
        return False

    if not os.path.exists(file_path):
        logger.error(f"Excel-файл не найден: {file_path}")
        return False

    logger.info(f"Начало импорта всех данных из Excel через xlwings: {file_path}")

    # --- ИСПРАВЛЕНО: Инициализируем переменные до try ---
    app = None
    wb = None
    # --- КОНЕЦ ИСПРАВЛЕНИЯ ---

    try:
        # Открываем Excel-приложение в скрытом режиме
        app = xw.App(visible=False)
        # --- ИСПРАВЛЕНО: Убран visible=False из books.open ---
        wb = app.books.open(
            file_path,
            update_links=False,  # Не обновлять внешние ссылки
            read_only=True       # Открыть только для чтения
            # visible=False УДАЛЕН
        )
        # --- КОНЕЦ ИСПРАВЛЕНИЯ ---
        logger.debug(f"Книга '{file_path}' открыта через xlwings (read-only).")

        # --- НОВОЕ: Сообщаем о начале ---
        if progress_callback:
            progress_callback(0, f"Открыт файл: {os.path.basename(file_path)}")
        # --- КОНЕЦ НОВОГО ---

        total_sheets = len(wb.sheets)
        processed_sheets = 0

        # Импортируем листы
        for sheet in wb.sheets:
            processed_sheets += 1
            logger.info(f"Обработка листа: {sheet.name}")
            
            # --- НОВОЕ: Обновляем прогресс ---
            if progress_callback:
                percent = int((processed_sheets / total_sheets) * 50) # Первые 50% — листы
                progress_callback(percent, f"Обработка листа: {sheet.name}")
            # --- КОНЕЦ НОВОГО ---
            
            # Сохраняем информацию о листе
            sheet_id = storage.save_sheet(project_id=1, sheet_name=sheet.name)
            if sheet_id is None:
                logger.error(f"Не удалось сохранить лист '{sheet.name}'.")
                return False

            # --- ИМПОРТ СЫРЫХ ДАННЫХ, ФОРМУЛ И ФОРМАТОВ ---
            raw_data_list, formulas_list, formats_list = _extract_raw_formula_and_format_data_from_sheet(sheet, progress_callback, processed_sheets, total_sheets)

            if raw_data_list:
                if not storage.save_sheet_raw_data(sheet.name, raw_data_list):
                    logger.error(f"Не удалось сохранить 'сырые данные' для листа '{sheet.name}'.")
                    return False
                logger.debug(f"Сохранено {len(raw_data_list)} записей 'сырых данных' для листа '{sheet.name}'.")

            if formulas_list:
                if not storage.save_sheet_formulas(sheet_id, formulas_list):
                    logger.error(f"Не удалось сохранить 'формулы' для листа '{sheet.name}'.")
                    return False
                logger.debug(f"Сохранено {len(formulas_list)} записей 'формул' для листа '{sheet.name}'.")

            if formats_list:
                # storage.save_sheet_formats — такой функции может не быть
                # Пока логируем, что форматы извлечены
                logger.debug(f"Извлечено {len(formats_list)} записей 'числовых форматов' для листа '{sheet.name}'.")
                # TODO: Реализовать сохранение форматов в БД, если поддерживается
                # Пример: storage.save_sheet_formats(sheet_id, formats_list)

            # --- ИМПОРТ СТИЛЕЙ ---
            # ВРЕМЕННО ОТКЛЮЧЕНО ДЛЯ УЛУЧШЕНИЯ ПРОИЗВОДИТЕЛЬНОСТИ
            # styles_list = _extract_styles_from_xlwings_sheet(sheet, sheet_id)
            # if styles_list:
            #     if not storage.save_sheet_styles(sheet_id, styles_list):
            #         logger.error(f"Не удалось сохранить 'стили' для листа '{sheet.name}'.")
            #         return False
            #     logger.debug(f"Сохранено {len(styles_list)} записей 'стилей' для листа '{sheet.name}'.")
            # --- КОНЕЦ ОТКЛЮЧЕНИЯ ---

            # --- ИМПОРТ ДИАГРАММ ---
            charts_list = _extract_charts_from_xlwings_sheet(sheet, sheet_id)
            if charts_list:
                if not storage.save_sheet_charts(sheet_id, charts_list):
                    logger.error(f"Не удалось сохранить 'диаграммы' для листа '{sheet.name}'.")
                    return False
                logger.debug(f"Сохранено {len(charts_list)} записей 'диаграмм' для листа '{sheet.name}'.")

            # --- ИМПОРТ ОБЪЕДИНЁННЫХ ЯЧЕЕК ---
            merged_cells_list = _extract_merged_cells_from_xlwings_sheet(sheet, sheet_id)
            if merged_cells_list:
                if not storage.save_sheet_merged_cells(sheet_id, merged_cells_list):
                    logger.error(f"Не удалось сохранить 'объединённые ячейки' для листа '{sheet.name}'.")
                    return False
                logger.debug(f"Сохранено {len(merged_cells_list)} записей 'объединённых ячеек' для листа '{sheet.name}'.")

            # --- ИМПОРТ ИМЕНОВАННЫХ ДИАПАЗОНОВ ---
            # wb.names может содержать именованные диапазоны
            # TODO: Реализовать извлечение и сохранение именованных диапазонов
            logger.debug(f"Именованные диапазоны на листе '{sheet.name}' пока не обрабатываются xlwings-импортером.")

            # --- ИМПОРТ МЕТАДАННЫХ ЛИСТА ---
            # sheet.visible, sheet.name, sheet.index и т.д.
            # TODO: Реализовать сохранение метаданных листа
            logger.debug(f"Метаданные листа '{sheet.name}' пока не обрабатываются xlwings-импортером.")

        logger.info(f"Обработка всех листов в '{file_path}' завершена.")

        # --- НОВОЕ: ИМПОРТ МЕТАДАННЫХ ---
        metadata_dict = _extract_metadata_from_xlwings_workbook(wb)
        if metadata_dict:
            # Предполагаем project_id = 1 для MVP
            project_id = 1
            if not storage.save_project_metadata(project_id, metadata_dict):
                logger.error("Не удалось сохранить метаданные проекта.")
            else:
                logger.info(f"Сохранено {len(metadata_dict)} записей метаданных.")
        # --- КОНЕЦ НОВОГО ---

        # --- НОВОЕ: Сообщаем о завершении ---
        if progress_callback:
            progress_callback(100, "Импорт завершён")
        # --- КОНЕЦ НОВОГО ---

        # Всё прошло успешно
        return True

    except Exception as e:
        logger.error(f"Ошибка при импорте через xlwings из '{file_path}': {e}", exc_info=True)
        return False

    finally:
        # --- ИСПРАВЛЕНО: Проверяем, были ли инициализированы app и wb ---
        if wb:
            try:
                wb.close()
                logger.debug("Книга xlwings закрыта.")
            except Exception as e:
                logger.warning(f"Ошибка при закрытии книги xlwings: {e}")
        if app:
            try:
                app.quit()
                logger.debug("Приложение xlwings закрыто.")
            except Exception as e:
                logger.warning(f"Ошибка при закрытии приложения xlwings: {e}")
        # --- КОНЕЦ ИСПРАВЛЕНИЯ ---

# --- Вспомогательные функции для преобразования данных xlwings в формат storage ---

def _extract_raw_formula_and_format_data_from_sheet(
    sheet: xw.Sheet,
    progress_callback: Optional[Callable[[int, str], None]] = None,
    sheet_index: int = 0,
    total_sheets: int = 1
) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]: # <-- ИЗМЕНЕНО
    """
    Извлекает raw_data, formulas и formats из листа xlwings.

    Args:
        sheet (xw.Sheet): Лист xlwings.
        progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
        sheet_index (int): Индекс текущего листа (для расчёта процента).
        total_sheets (int): Общее количество листов.

    Returns:
        tuple: (raw_data_list, formulas_list, formats_list)
    """
    raw_data_list = []
    formulas_list = []
    formats_list = []

    used_range = sheet.used_range
    if not used_range:
        return raw_data_list, formulas_list, formats_list

    values_matrix = used_range.value
    formulas_matrix = used_range.formula

    start_row, start_col = used_range.row, used_range.column
    n_rows, n_cols = used_range.rows.count, used_range.columns.count

    total_cells = n_rows * n_cols
    processed_cells = 0

    # xlwings использует 1-based индексацию
    for i in range(n_rows):
        for j in range(n_cols):
            value = values_matrix[i][j] if values_matrix else None
            formula = formulas_matrix[i][j] if formulas_matrix else None
            # Получаем адрес ячейки
            cell_address = sheet.range(start_row + i, start_col + j).address

            if value is not None:
                raw_data_list.append({
                    "cell_address": cell_address,
                    "value": value
                })
            
            if formula and isinstance(formula, str) and formula.startswith('='):
                formulas_list.append({
                    "cell_address": cell_address,
                    "formula": formula
                })

            # --- ИЗВЛЕЧЕНИЕ ЧИСЛОВОГО ФОРМАТА ---
            cell_xl = sheet.range(start_row + i, start_col + j).api  # COM-объект Excel
            number_format = cell_xl.NumberFormat
            if number_format:
                formats_list.append({
                    "cell_address": cell_address,
                    "number_format": str(number_format)
                })

            processed_cells += 1
            # --- НОВОЕ: Обновляем прогресс каждые 10 ячеек ---
            if progress_callback and processed_cells % 10 == 0: # <-- ИЗМЕНЕНО С 1000 НА 10
                # Рассчитываем общий процент: первые 50% — листы, следующие 50% — ячейки
                sheet_percent = (sheet_index / total_sheets) * 50
                cell_percent = (processed_cells / total_cells) * 50 if total_cells > 0 else 0
                total_percent = int(sheet_percent + cell_percent)
                progress_callback(min(total_percent, 99), f"Обработка ячеек листа {sheet.name} ({processed_cells}/{total_cells})")
            # --- КОНЕЦ НОВОГО ---

    return raw_data_list, formulas_list, formats_list

def _extract_styles_from_xlwings_sheet(sheet: xw.Sheet, sheet_id: int) -> List[Dict[str, Any]]:
    """
    Извлекает стили из листа xlwings.

    Args:
        sheet (xw.Sheet): Лист xlwings.
        sheet_id (int): ID листа в БД.

    Returns:
        List[Dict[str, Any]]: Список данных стилей.
    """
    styles_list = []

    used_range = sheet.used_range
    if not used_range:
        return styles_list

    start_row, start_col = used_range.row, used_range.column
    n_rows, n_cols = used_range.rows.count, used_range.columns.count

    for i in range(n_rows):
        for j in range(n_cols):
            cell = sheet.range(start_row + i, start_col + j)
            cell_address = cell.address
            xl_api_obj = cell.api

            style_dict = _serialize_style_from_xlwings_range_api(xl_api_obj)
            if style_dict:
                style_json = json.dumps(style_dict, ensure_ascii=False)
                styles_list.append({
                    "range_address": cell_address,
                    "style_attributes": style_json
                })

    return styles_list
def _serialize_style_from_xlwings_range_api(xl_range_api):
    """
    Сериализует стили COM-объекта xlwings Range.api в словарь.
    """
    style_dict = {}
    try:
        font = xl_range_api.Font
        style_dict['font'] = {
            'name': font.Name,
            'size': float(font.Size) if font.Size else None,
            'bold': bool(font.Bold),
            'italic': bool(font.Italic),
            'color': int(font.Color) if font.Color else None,  # Excel Color — это число
        }

        interior = xl_range_api.Interior
        style_dict['interior'] = {
            'color': int(interior.Color) if interior.Color else None,
            'pattern': interior.Pattern,
        }

        borders = xl_range_api.Borders
        style_dict['borders'] = {
            'style': borders.LineStyle,
            'color': int(borders.Color) if borders.Color else None,
        }

        style_dict['number_format'] = xl_range_api.NumberFormat
        style_dict['horizontal_alignment'] = xl_range_api.HorizontalAlignment
        style_dict['vertical_alignment'] = xl_range_api.VerticalAlignment

    except Exception as e:
        logger.warning(f"Ошибка при сериализации стиля xlwings: {e}")
        return {}

    return style_dict
def _extract_charts_from_xlwings_sheet(sheet: xw.Sheet, sheet_id: int) -> List[Dict[str, Any]]:
    """
    Извлекает диаграммы из листа xlwings.

    Args:
        sheet (xw.Sheet): Лист xlwings.
        sheet_id (int): ID листа в БД.

    Returns:
        List[Dict[str, Any]]: Список данных диаграмм.
    """
    charts_list = []
    for chart_obj in sheet.charts:
        # Пока сохраняем только имя, как заглушка
        # В будущем можно расширить для извлечения данных диаграммы
        chart_data = {
            "chart_name": chart_obj.name,
            "sheet_id": sheet_id
        }
        charts_list.append(chart_data)
    return charts_list
def _extract_merged_cells_from_xlwings_sheet(sheet: xw.Sheet, sheet_id: int) -> List[str]:
    """
    Извлекает объединённые ячейки из листа xlwings.

    Args:
        sheet (xw.Sheet): Лист xlwings.
        sheet_id (int): ID листа в БД.

    Returns:
        List[str]: Список строк адресов диапазонов (например, ['A1:B2', 'C3:D4']).
    """
    # Простой обход used_range для поиска объединённых ячеек
    used_range = sheet.used_range
    if not used_range:
        return []

    merged_set = set()
    start_row, start_col = used_range.row, used_range.column
    n_rows, n_cols = used_range.rows.count, used_range.columns.count

    for i in range(n_rows):
        for j in range(n_cols):
            cell = sheet.range(start_row + i, start_col + j)
            # Если ячейка — часть объединённой области
            if cell.merge_area.count > 1:
                # Получаем адрес всей объединённой области
                merged_addr = cell.merge_area.address
                merged_set.add(merged_addr)

    merged_list = list(merged_set)
    return merged_list

# --- НОВОЕ: ИЗВЛЕЧЕНИЕ МЕТАДАННЫХ ---
def _extract_metadata_from_xlwings_workbook(wb: xw.Book) -> Dict[str, Any]:
    """
    Извлекает метаданные из книги Excel через xlwings.

    Args:
        wb (xw.Book): Книга xlwings.

    Returns:
        Dict[str, Any]: Словарь с ключами и значениями метаданных.
    """
    metadata = {}
    try:
        # Получаем COM-объект книги
        wb_api = wb.api

        # Получаем встроенные свойства документа
        builtin_props = wb_api.BuiltinDocumentProperties

        # Список интересующих нас свойств
        prop_names = [
            "Title", "Subject", "Author", "Keywords", "Comments",
            "Template", "Last Author", "Revision Number", "Application Name",
            "Last Print Date", "Creation Date", "Last Save Time",
            "Total Edit Time", "Number of Pages", "Number of Words",
            "Number of Characters", "Security", "Category", "Format",
            "Manager", "Company", "Bytes", "Lines", "Paragraphs",
            "Slides", "Notes", "Hidden Slides", "MM Clips",
            "Scale Crop", "Heading Pairs", "Titles of Parts",
            "Links up-to-date", "Characters with Spaces", "Shared Doc",
            "Hyperlink Base", "HRefs", "Bookmarks", "Target Frame",
            "Encoding", "Decryption Status", "Permission Status",
            "Content Status", "Language", "Version Info", "Digital Signature",
            "Lock Comments", "Web View", "Server URL", "Share Point URL",
            "Unique Identifier", "Original Filename", "Original Path",
            "Last Modified By", "Content Type", "Content Created",
            "Date Last Printed", "Date Created", "Date Modified",
            "Byte Count", "Line Count", "Paragraph Count", "Slide Count",
            "Note Count", "Hidden Slide Count", "Multimedia Clip Count",
            "Scale Crop", "Heading Pair Count", "Title of Part Count",
            "Character Count With Spaces", "Shared Document",
            "Hyperlink Base", "Hyperlinks Changed", "Digital Signature",
            "Encryption Type", "Password Encryption Provider",
            "Password Encryption Algorithm", "Password Encryption Key Length",
            "Password Encryption Hash Value", "Password Encryption Salt Value",
            "Password Encryption Spin Count", "Document Integrity",
            "ContentTypeId", "Document Management Policy",
            "Document Security Store", "Document Workspace",
            "Publishing Path", "Sharing Capability", "Sharing Permission",
            "Sharing Scope", "Viewer Type", "Sensitivity Label",
            "Sensitivity Label Extension", "Sensitivity Metadata",
            "Thumbnail", "Thumbnail Large", "Thumbnail Small",
            "Thumbnail Extra Large", "Thumbnail Extra Small",
            "Thumbnail Custom", "Thumbnail Custom Large",
            "Thumbnail Custom Small", "Thumbnail Custom Extra Large",
            "Thumbnail Custom Extra Small"
        ]

        for name in prop_names:
            try:
                prop = builtin_props(name)
                value = prop.Value
                if value is not None:
                    # Преобразуем даты в строки, если нужно
                    if isinstance(value, (int, float)) and name in ["Creation Date", "Last Save Time", "Content Created", "Date Created", "Date Modified", "Date Last Printed"]:
                        # Excel хранит даты как числа (OLE Automation Date)
                        import datetime
                        try:
                            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)
                            value = dt.isoformat()
                        except:
                            pass
                    metadata[name] = str(value)
            except:
                # Свойство не задано или ошибка доступа
                pass

        # Также можно получить имя файла
        try:
            metadata["File Name"] = wb.name
        except:
            pass

        # И путь к файлу
        try:
            metadata["Full Path"] = wb.fullname
        except:
            pass

    except Exception as e:
        logger.warning(f"Ошибка при извлечении метаданных xlwings: {e}")

    return metadata
# --- КОНЕЦ НОВОГО ---

# TODO: Реализовать остальные функции:
# _extract_named_ranges_from_xlwings_workbook(wb_obj) -> List[Dict[str, Any]]
# и т.д.
