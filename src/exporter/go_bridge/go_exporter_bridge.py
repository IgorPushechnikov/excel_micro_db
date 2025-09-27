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
from pathlib import Path
from typing import Any, Optional

from src.exporter.go_bridge.export_data_model import ExportData, SheetData, Formula, Style, Chart, ChartSeries, ProjectMetadata
from src.storage.base import ProjectDBStorage


logger = logging.getLogger(__name__)


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

        # 1. Загрузка редактируемых данных
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
        data_matrix = [[None for _ in range(max_col + 1)] for _ in range(max_row + 1)]

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
            # Предполагается, что chart_item - это словарь с ключами, соответствующими Chart и ChartSeries
            # Нужно аккуратно преобразовать его в Pydantic-модели.
            try:
                # Если данные диаграммы уже в правильном формате, можно использовать model_validate
                chart = Chart.model_validate(chart_item)
                charts_list.append(chart)
            except Exception as e:
                logger.warning(f"Не удалось преобразовать данные диаграммы для листа '{sheet_name}': {e}")
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

    def export_to_xlsx(self, output_path: Path, temp_dir: Path = None) -> bool:
        """
        Выполняет экспорт данных в XLSX-файл с помощью Go-утилиты.

        Args:
            output_path: Путь для сохранения итогового XLSX-файла.
            temp_dir: Временная директория для хранения JSON-файла. Если None, используется системная.

        Returns:
            bool: True в случае успеха, False в случае ошибки.
        """
        try:
            # 1. Подготовка данных
            logger.info("Preparing export data...")
            export_data = self._prepare_export_data()

            # 2. Сериализация в JSON
            temp_json_path = (temp_dir or Path.cwd()) / "export_data.json"
            with open(temp_json_path, 'w', encoding='utf-8') as f:
                json.dump(export_data.model_dump(), f, ensure_ascii=False, indent=2)
            logger.info(f"Export data serialized to {temp_json_path}")

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
                # Удаляем временный JSON-файл
                temp_json_path.unlink(missing_ok=True)
                return True
            else:
                logger.error(f"Go exporter failed with return code {result.returncode}")
                logger.error(f"Stderr: {result.stderr}")
                logger.error(f"Stdout: {result.stdout}")
                return False

        except Exception as e:
            logger.exception(f"An error occurred during export: {e}")
            return False