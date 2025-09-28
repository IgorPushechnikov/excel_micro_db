"""
Мост между основным Python-приложением и Go-утилитой для экспорта.

Этот модуль отвечает за:
1. Сбор данных из ProjectDBStorage.
2. Преобразование их в валидированную структуру ExportData.
3. Сериализацию в JSON-файл.
4. Вызов Go-утилиты и обработку её результата.
"""

import json
import subprocess
import logging
import tempfile
from pathlib import Path
from typing import Any, List, Optional, Dict, Tuple

from src.exporter.go_bridge.export_data_model import ExportData, SheetData, Formula, Style, Chart, ChartSeries, ProjectMetadata
from src.storage.base import ProjectDBStorage


logger = logging.getLogger(__name__)


def _convert_chart_position_db_to_address(position_data: Dict[str, Any]) -> str:
    """
    Преобразует словарь позиции диаграммы из БД в строку адреса ячейки (например, "A1").
    Использует 'from_col' и 'from_row' из данных позиции.

    Args:
        position_data (Dict[str, Any]): Словарь с данными позиции из БД.

    Returns:
        str: Строка адреса ячейки (например, "A1") или "A1" по умолчанию при ошибке.
    """
    try:
        # Извлекаем индексы столбца и строки (0-based)
        col_index = position_data.get('from_col', 0)
        row_index = position_data.get('from_row', 0)
        
        # Преобразуем 0-based индекс в 1-based номер для Excel
        row_number = row_index + 1
        
        # Преобразуем 0-based индекс столбца в имя столбца Excel (A, B, ..., Z, AA, AB, ...)
        # Алгоритм: https://stackoverflow.com/questions/2386196/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        column_name = ""
        temp_col_index = col_index
        while temp_col_index >= 0:
            column_name = chr(temp_col_index % 26 + ord('A')) + column_name
            temp_col_index = temp_col_index // 26 - 1
        
        # Если column_name осталась пустой (например, при col_index < 0), используем A
        if not column_name:
            column_name = "A"
            logger.warning(f"Имя столбца диаграммы оказалось пустым для индекса {col_index}. Используется 'A'.")
        
        return f"{column_name}{row_number}"
    except Exception as e:
        logger.warning(f"Ошибка при преобразовании позиции диаграммы {position_data}: {e}. Используется A1.")
        return "A1" # Возвращаем значение по умолчанию в случае ошибки


class GoExporterBridge:
    """Класс-мост для взаимодействия с Go-утилитой экспорта."""

    def __init__(self, db_storage: ProjectDBStorage, go_exporter_path: Optional[str] = None):
        """
        Инициализация моста.

        Args:
            db_storage: Экземпляр ProjectDBStorage для доступа к данным.
            go_exporter_path: Путь к исполняемому файлу Go-утилиты (go_excel_exporter.exe).
                              Если не указан, используется путь по умолчанию.
        """
        self.db_storage = db_storage
        if go_exporter_path is None:
            # Путь по умолчанию относительно этого файла
            default_path = Path(__file__).parent.parent / "go" / "go_excel_exporter.exe"
            self.go_exporter_path = default_path
        else:
            self.go_exporter_path = Path(go_exporter_path)
        
        if not self.go_exporter_path.exists():
            raise FileNotFoundError(f"Go exporter not found at: {self.go_exporter_path}")

    def _prepare_sheet_data(self, sheet_id: int, sheet_name: str) -> SheetData:
        """
        Подготавливает данные для одного листа из БД.

        Args:
            sheet_id: ID листа в БД.
            sheet_name: Имя листа.

        Returns:
            SheetData: Подготовленные данные для листа.
        """
        logger.debug(f"Подготовка данных для листа '{sheet_name}' (ID: {sheet_id})...")

        # 1. Загрузка "сырых" (редактируемых) данных
        editable_data_rows = self.db_storage.load_sheet_editable_data(sheet_id, sheet_name)
        # editable_data_rows - это список словарей [{'cell_address': 'A1', 'value': 'Header'}, ...]
        # Нам нужно преобразовать это в двумерный список (матрицу)

        # Для этого сначала определим максимальные индексы строк и столбцов
        max_row = 0
        max_col = 0
        cell_dict = {}  # {(row, col): value}

        for item in editable_data_rows:
            address = item['cell_address']
            value = item['value']
            # Преобразуем адрес Excel (например, 'AB123') в индексы (row, col)
            row, col = self._excel_address_to_indices(address)
            if row is not None and col is not None:
                max_row = max(max_row, row)
                max_col = max(max_col, col)
                cell_dict[(row, col)] = str(value) if value is not None else None

        # Создаем пустую матрицу данных
        # Явно указываем тип, чтобы избежать ошибок типизации
        data_matrix: List[List[Optional[str]]] = [[None for _ in range(max_col + 1)] for _ in range(max_row + 1)]

        # Заполняем матрицу значениями из словаря
        for (row, col), value in cell_dict.items():
            data_matrix[row][col] = value

        # 2. Загрузка формул
        formulas_data = self.db_storage.load_sheet_formulas(sheet_id)
        formulas_list = [
            Formula(cell=item['cell_address'], formula=item['formula'])
            for item in formulas_data
        ]

        # 3. Загрузка стилей
        # TODO: Реализовать преобразование стилей из формата БД в формат, понятный Go-экспортеру.
        # Пока оставим пустым списком.
        styles_list = []

        # 4. Загрузка диаграмм
        charts_data = self.db_storage.load_sheet_charts(sheet_id)
        charts_list = []
        for chart_item in charts_data:
            # chart_item - это словарь вида {'chart_data': '...json string...'}
            # Нужно извлечь и распарсить JSON-строку
            chart_data_str = chart_item.get('chart_data', '{}')
            if not chart_data_str:
                continue
            try:
                # Десериализуем JSON-строку в словарь Python
                chart_dict_from_db = json.loads(chart_data_str)
                
                # --- Адаптация данных из БД в формат Chart ---
                adapted_chart_dict = {}
                
                # 1. type -> type
                adapted_chart_dict['type'] = chart_dict_from_db.get('type', 'col')
                
                # 2. position (сложный объект) -> position (строка)
                position_data = chart_dict_from_db.get('position', {})
                if isinstance(position_data, dict):
                    adapted_chart_dict['position'] = _convert_chart_position_db_to_address(position_data)
                else:
                    # Если position не словарь, используем значение по умолчанию
                    logger.warning(f"Поле 'position' диаграммы не является словарем: {position_data}. Используется A1.")
                    adapted_chart_dict['position'] = "A1"
                
                # 3. title -> title (опционально)
                title = chart_dict_from_db.get('title')
                if title is not None:
                    adapted_chart_dict['title'] = str(title)
                # Если title отсутствует, Chart.model_validate установит значение по умолчанию (None)
                
                # 4. series -> series (нужно адаптировать ChartSeries тоже)
                series_data_list = chart_dict_from_db.get('series', [])
                adapted_series_list = []
                if isinstance(series_data_list, list):
                    for series_item in series_data_list:
                        if isinstance(series_item, dict):
                            adapted_series_item = {}
                            # val_range -> values (обязательно)
                            val_range = series_item.get('val_range')
                            if val_range:
                                adapted_series_item['values'] = val_range
                            else:
                                # Если val_range отсутствует, пропускаем этот элемент series
                                logger.warning(f"Элемент series не содержит обязательное поле 'val_range': {series_item}")
                                continue
                            
                            # cat_range -> categories (опционально)
                            cat_range = series_item.get('cat_range')
                            if cat_range is not None: # Может быть пустая строка, которая тоже валидна
                                adapted_series_item['categories'] = cat_range
                            else:
                                # Если cat_range отсутствует, явно устанавливаем None
                                adapted_series_item['categories'] = None
                            
                            # name -> name (опционально)
                            name = series_item.get('name')
                            if name is not None: # Может быть пустая строка, которая тоже валидна
                                adapted_series_item['name'] = str(name)
                            else:
                                # Если name отсутствует, явно устанавливаем None
                                adapted_series_item['name'] = None
                            
                            # Создаем объект ChartSeries из адаптированного словаря
                            try:
                                chart_series = ChartSeries.model_validate(adapted_series_item)
                                adapted_series_list.append(chart_series)
                            except Exception as e:
                                logger.warning(f"Не удалось создать ChartSeries из адаптированного элемента: {e}", exc_info=True)
                                continue
                        else:
                            logger.warning(f"Элемент series не является словарем: {series_item}")
                else:
                    logger.warning(f"Поле 'series' диаграммы не является списком: {series_data_list}")
                
                adapted_chart_dict['series'] = adapted_series_list
                
                # --- Конец адаптации ---
                
                # Теперь валидируем адаптированный словарь с помощью Pydantic
                chart = Chart.model_validate(adapted_chart_dict)
                charts_list.append(chart)
            except json.JSONDecodeError as je:
                logger.warning(f"Не удалось распарсить JSON диаграммы для листа '{sheet_name}': {je}")
                continue
            except Exception as e:
                logger.warning(f"Не удалось преобразовать данные диаграммы для листа '{sheet_name}': {e}", exc_info=True)
                continue

        return SheetData(
            name=sheet_name,
            data=data_matrix,
            formulas=formulas_list,
            styles=styles_list,
            charts=charts_list
        )

    def _excel_address_to_indices(self, address: str) -> tuple[int, int] | tuple[None, None]:
        """
        Преобразует адрес ячейки Excel (например, 'AB123') в индексы (row, col).
        Индексы 0-базированные.

        Args:
            address: Адрес ячейки в формате Excel.

        Returns:
            tuple[int, int]: (row_index, col_index) или (None, None) в случае ошибки.
        """
        import re
        match = re.match(r"([A-Za-z]+)(\d+)", address)
        if not match:
            logger.error(f"Неверный формат адреса ячейки: {address}")
            return (None, None)

        col_str, row_str = match.groups()
        row_index = int(row_str) - 1  # Excel использует 1-индексацию для строк

        # Преобразуем имя столбца (A, B, ..., Z, AA, AB, ...) в индекс
        col_index = 0
        for char in col_str.upper():
            col_index = col_index * 26 + (ord(char) - ord('A') + 1)
        col_index -= 1  # 0-базированный индекс

        return (row_index, col_index)

    def _prepare_export_data(self) -> ExportData:
        """
        Подготавливает полную структуру данных для экспорта.

        Returns:
            ExportData: Валидированные данные для передачи в Go.
        """
        logger.info("Подготовка полных данных для экспорта...")

        # 1. Получаем список всех листов из БД
        sheets_metadata = self.db_storage.load_all_sheets_metadata(project_id=1)  # Предполагаем project_id=1 для MVP
        if not sheets_metadata:
            logger.warning("Не найдено ни одного листа для экспорта.")
            # Создаем пустой лист, чтобы Go-экспортер не падал
            sheets_metadata = [{'sheet_id': 1, 'name': 'Sheet1'}]

        sheets_data = []
        for sheet_info in sheets_metadata:
            sheet_id = sheet_info['sheet_id']
            sheet_name = sheet_info['name']
            try:
                sheet_data = self._prepare_sheet_data(sheet_id, sheet_name)
                sheets_data.append(sheet_data)
            except Exception as e:
                logger.error(f"Ошибка при подготовке данных для листа '{sheet_name}' (ID: {sheet_id}): {e}", exc_info=True)
                # Пропускаем проблемный лист, но продолжаем экспорт остальных
                continue

        # 2. Получаем метаданные проекта из БД
        # TODO: Реализовать загрузку метаданных проекта (project_name, author, created_at) из БД.
        # Пока используем заглушки.
        project_metadata = self.db_storage.load_sheet_metadata("__project__") # Пример, как может храниться метадата проекта
        if project_metadata:
            project_name = project_metadata.get("project_name", "Excel Micro DB Project")
            author = project_metadata.get("author", "Unknown")
            created_at = project_metadata.get("created_at", "2025-01-01T00:00:00Z")
        else:
            # Используем значения по умолчанию
            project_name = "Excel Micro DB Project"
            author = "User"
            created_at = "2025-01-01T00:00:00Z"

        metadata = ProjectMetadata(
            project_name=project_name,
            author=author,
            created_at=created_at
        )

        return ExportData(metadata=metadata, sheets=sheets_data)

    def export_to_xlsx(self, output_path: Path, temp_dir: Optional[Path] = None, debug_json_path: Optional[Path] = None) -> bool:
        """
        Выполняет экспорт данных в XLSX-файл с помощью Go-утилиты.

        Args:
            output_path: Путь для сохранения итогового XLSX-файла.
            temp_dir: Временная директория для хранения JSON-файла. Если None, используется временная директория ОС.
            debug_json_path: Необязательный путь для сохранения отладочного JSON-файла.
                             Если указан, JSON будет сохранен по этому пути вместо временного файла.

        Returns:
            bool: True в случае успеха, False в случае ошибки.
        """
        temp_json_path = None
        try:
            # 1. Подготовка данных
            logger.info("Preparing export data...")
            export_data = self._prepare_export_data()

            # 2. Сериализация в JSON
            if debug_json_path:
                # Если указан debug_json_path, используем его
                temp_json_path = debug_json_path
                # Создаем родительские директории, если нужно
                temp_json_path.parent.mkdir(parents=True, exist_ok=True)
                with open(temp_json_path, 'w', encoding='utf-8') as f:
                    json.dump(export_data.model_dump(), f, ensure_ascii=False, indent=2)
                logger.info(f"Export data serialized to debug JSON file: {temp_json_path}")
            elif temp_dir is not None:
                temp_json_path = temp_dir / "export_data.json"
                with open(temp_json_path, 'w', encoding='utf-8') as f:
                    json.dump(export_data.model_dump(), f, ensure_ascii=False, indent=2)
                logger.info(f"Export data serialized to {temp_json_path}")
            else:
                # Создаем временный файл, который автоматически удалится (если debug_json_path не указан)
                with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as tmp_file:
                    json.dump(export_data.model_dump(), tmp_file, ensure_ascii=False, indent=2)
                    temp_json_path = Path(tmp_file.name)
                logger.info(f"Export data serialized to temporary file: {temp_json_path}")

            # 3. Вызов Go-утилиты
            logger.info("Calling Go exporter...")
            result = subprocess.run([
                str(self.go_exporter_path),
                "-input", str(temp_json_path),
                "-output", str(output_path)
            ], capture_output=True, text=True)

            # 4. Обработка результата
            if result.returncode == 0:
                logger.info(f"Export successful! File saved to {output_path}")
                success = True
            else:
                logger.error(f"Go exporter failed with return code {result.returncode}")
                logger.error(f"Stderr: {result.stderr}")
                logger.error(f"Stdout: {result.stdout}")
                success = False
            
            return success

        except Exception as e:
            logger.exception(f"An error occurred during export: {e}")
            return False
        finally:
            # Удаляем временный JSON-файл в любом случае, если debug_json_path не был указан
            if temp_json_path and not debug_json_path:
                try:
                    temp_json_path.unlink(missing_ok=True)
                    logger.debug(f"Temporary JSON file deleted: {temp_json_path}")
                except Exception as e:
                    logger.warning(f"Could not delete temporary file {temp_json_path}: {e}")