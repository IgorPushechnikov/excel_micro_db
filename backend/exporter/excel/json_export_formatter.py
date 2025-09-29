import json
import os
import logging
from typing import Dict, Any, List, Optional, Union
from datetime import datetime

# Импортируем openpyxl для работы с адресами ячеек
from openpyxl.utils import coordinate_to_tuple, get_column_letter

# Импортируем хранилище для получения данных
from src.storage.base import ProjectDBStorage

# Получаем логгер
logger = logging.getLogger(__name__)


def _convert_none_to_null(obj):
    """
    Рекурсивно преобразует None в null для корректной сериализации в JSON.
    В Python `None` сериализуется в `null` по умолчанию, но если в данных
    есть другие объекты (например, из numpy), это может понадобиться.
    Пока оставим как есть, так как стандартный json.dumps обрабатывает None.
    """
    return obj


def export_project_to_json_format(project_db_path: str, output_json_path: str) -> bool:
    """
    Экспортирует данные проекта из БД в формат JSON, пригодный для Go-экспортёра.

    Args:
        project_db_path (str): Путь к файлу БД проекта (project_data.db).
        output_json_path (str): Путь к выходному JSON-файлу.

    Returns:
        bool: True, если экспорт успешен, иначе False.
    """
    logger.info(f"Начало экспорта проекта в JSON-формат для Go-экспортёра: {output_json_path}")
    
    if not os.path.exists(project_db_path):
        logger.error(f"Файл БД проекта не найден: {project_db_path}")
        return False

    storage = None
    try:
        # 1. Подключение к БД проекта
        storage = ProjectDBStorage(project_db_path)
        if not storage.connect():
            logger.error(f"Не удалось подключиться к БД проекта: {project_db_path}")
            return False
        
        # 2. Подготовка структуры ExportData
        export_data: Dict[str, Any] = {
            "metadata": {},
            "sheets": []
        }
        
        # --- Заполнение метаданных (заглушка, можно улучшить) ---
        # TODO: Получить реальные метаданные проекта из БД или другого источника
        export_data["metadata"] = {
            "project_name": os.path.basename(os.path.dirname(project_db_path)), # Имя папки проекта
            "author": "Unknown", # Заглушка
            "created_at": datetime.now().isoformat() # Текущее время в формате ISO 8601
        }
        logger.debug(f"Метаданные проекта подготовлены: {export_data['metadata']}")
        
        # --- Получение списка листов из БД ---
        sheets_metadata = storage.load_all_sheets_metadata()
        logger.debug(f"Получен список листов из БД: {sheets_metadata}")
        
        if not sheets_metadata:
            logger.warning("В проекте не найдено листов для экспорта в JSON.")
        
        # --- Итерация по листам и заполнение данных ---
        for sheet_info in sheets_metadata:
            sheet_id = sheet_info['sheet_id']
            sheet_name = sheet_info['name']
            logger.info(f"Экспорт данных для листа: '{sheet_name}' (ID: {sheet_id})")
            
            sheet_data_item: Dict[str, Any] = {
                "name": sheet_name,
                "data": [],
                "formulas": [],
                "styles": [], # Будет заполнен ниже
                "charts": []  # Будет заполнен ниже
            }
            
            # --- Загрузка "сырых данных" ---
            raw_data_records = storage.load_sheet_raw_data(sheet_name)
            logger.debug(f"Загружено {len(raw_data_records)} записей 'сырых данных' для листа '{sheet_name}'.")
            
            # Простая логика: собрать все данные в словарь {address: value},
            # затем определить максимальные строку/столбец и создать двумерный массив.
            # Это не самый эффективный способ для больших таблиц, но подходит для MVP.
            cell_data_map = {}
            max_row = 0
            max_col = 0
            
            for record in raw_data_records:
                address = record.get('cell_address')
                value = record.get('value')
                # value_type = record.get('value_type') # Может понадобиться позже
                
                if address:
                    # Преобразуем адрес в координаты (например, A1 -> (1, 1))
                    # Нужна вспомогательная функция для этого
                    try:
                        from openpyxl.utils import coordinate_to_tuple
                        row, col = coordinate_to_tuple(address)
                        max_row = max(max_row, row)
                        max_col = max(max_col, col)
                        cell_data_map[address] = value
                    except Exception as e:
                        logger.warning(f"Не удалось преобразовать адрес ячейки '{address}': {e}")
            
            # Создаем двумерный массив данных (список списков)
            # Индексация в Python с 0, в Excel с 1. В JSON будем использовать 0-индексацию для массивов.
            if max_row > 0 and max_col > 0:
                sheet_data_array = []
                for r in range(1, max_row + 1): # Excel rows are 1-indexed
                    row_data = []
                    for c in range(1, max_col + 1): # Excel cols are 1-indexed
                        cell_address = f"{get_column_letter(c)}{r}"
                        cell_value = cell_data_map.get(cell_address)
                        # В Go-структуре *string означает указатель на строку или nil.
                        # В Python мы будем использовать None для пустых ячеек.
                        # json.dumps корректно сериализует None в null.
                        row_data.append(cell_value if cell_value is not None else None)
                    sheet_data_array.append(row_data)
                
                sheet_data_item["data"] = sheet_data_array
                logger.debug(f"Данные листа '{sheet_name}' собраны в двумерный массив размером {len(sheet_data_array)}x{len(sheet_data_array[0]) if sheet_data_array else 0}.")
            else:
                logger.debug(f"Лист '{sheet_name}' не содержит данных или не удалось определить размеры.")
                sheet_data_item["data"] = [] # Пустой массив
            
            # --- Загрузка формул ---
            formulas_records = storage.load_sheet_formulas(sheet_id)
            logger.debug(f"Загружено {len(formulas_records)} записей формул для листа '{sheet_name}' (ID: {sheet_id}).")
            
            for record in formulas_records:
                formula_item = {
                    "cell": record.get('cell_address', ''),
                    "formula": record.get('formula', '')
                }
                # Убедимся, что формула начинается с '='. Если нет, добавим.
                # Хотя, возможно, это уже сделано при сохранении в БД.
                if formula_item["formula"] and not formula_item["formula"].startswith('='):
                    formula_item["formula"] = '=' + formula_item["formula"]
                
                sheet_data_item["formulas"].append(formula_item)
            
            # --- Загрузка стилей ---
            styles_records = storage.load_sheet_styles(sheet_id)
            logger.debug(f"Загружено {len(styles_records)} записей стилей для листа '{sheet_name}' (ID: {sheet_id}).")
            
            for record in styles_records:
                # Предполагаем, что стиль хранится как JSON-строка в поле 'style_attributes'
                # и диапазон в поле 'range_address'
                style_attributes_str = record.get('style_attributes', '{}')
                range_address = record.get('range_address', '')
                
                try:
                    # Десериализуем JSON-строку атрибутов стиля обратно в словарь Python
                    style_attributes_dict = json.loads(style_attributes_str) if style_attributes_str else {}
                except json.JSONDecodeError as je:
                    logger.error(f"Ошибка разбора JSON стиля для листа '{sheet_name}', диапазон '{range_address}': {je}")
                    style_attributes_dict = {} # Используем пустой словарь в случае ошибки
                
                style_item = {
                    "range": range_address,
                    "style": style_attributes_dict # Передаем словарь напрямую
                }
                sheet_data_item["styles"].append(style_item)
            
            # --- Загрузка диаграмм ---
            charts_records = storage.load_sheet_charts(sheet_id)
            logger.debug(f"Загружено {len(charts_records)} записей диаграмм для листа '{sheet_name}' (ID: {sheet_id}).")
            
            for record in charts_records:
                # Предполагаем, что данные диаграммы хранятся как JSON-строка в поле 'chart_data'
                chart_data_str = record.get('chart_data')
                
                if chart_data_str:
                    try:
                        # Десериализуем JSON-строку, хранящуюся в БД, обратно в словарь
                        chart_data_dict = json.loads(chart_data_str)
                        
                        # Создаем элемент диаграммы для JSON, соответствующий Go-структуре
                        chart_item = {
                            "type": chart_data_dict.get('type', 'col'),
                            "position": chart_data_dict.get('position', 'A1'),
                            "title": chart_data_dict.get('title', ''),
                            "series": chart_data_dict.get('series', [])
                        }
                        
                        sheet_data_item["charts"].append(chart_item)
                        
                    except json.JSONDecodeError as je:
                        logger.error(f"Ошибка разбора JSON диаграммы для листа '{sheet_name}': {je}")
            
            # Добавляем собранные данные листа в основную структуру
            export_data["sheets"].append(sheet_data_item)
        
        # 3. Запись структуры ExportData в JSON-файл
        logger.debug(f"Подготовка к записи JSON-файла. Общее количество листов: {len(export_data['sheets'])}")
        
        # Убедимся, что директория для выходного файла существует
        output_dir = os.path.dirname(output_json_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        with open(output_json_path, 'w', encoding='utf-8') as f:
            # Используем ensure_ascii=False, чтобы кириллица и другие символы записывались корректно
            # Используем indent=2 для читаемости
            json.dump(export_data, f, ensure_ascii=False, indent=2, default=_convert_none_to_null)
        
        logger.info(f"Экспорт проекта в JSON-формат успешно завершен: {output_json_path}")
        return True
        
    except Exception as e:
        logger.error(f"Критическая ошибка при экспорте проекта в JSON: {e}", exc_info=True)
        return False
    finally:
        # 4. Закрытие соединения с БД
        if 'storage' in locals() and storage:
            storage.disconnect()
            logger.debug("Соединение с БД проекта закрыто после экспорта в JSON.")
