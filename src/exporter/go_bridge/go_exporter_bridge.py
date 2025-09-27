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
from typing import Any

from src.exporter.go_bridge.export_data_model import ExportData, SheetData, Formula, Style, Chart, ChartSeries, ProjectMetadata
from src.storage.base import ProjectDBStorage


logger = logging.getLogger(__name__)


class GoExporterBridge:
    """Класс-мост для взаимодействия с Go-утилитой экспорта."""

    def __init__(self, db_storage: ProjectDBStorage, go_exporter_path: str):
        """
        Инициализация моста.

        Args:
            db_storage: Экземпляр ProjectDBStorage для доступа к данным.
            go_exporter_path: Путь к исполняемому файлу Go-утилиты (go_excel_exporter.exe).
        """
        self.db_storage = db_storage
        self.go_exporter_path = Path(go_exporter_path)
        if not self.go_exporter_path.exists():
            raise FileNotFoundError(f"Go exporter not found at: {self.go_exporter_path}")

    def _prepare_sheet_data(self, sheet_name: str) -> SheetData:
        """
        Подготавливает данные для одного листа из БД.

        Этот метод должен быть реализован для извлечения:
        - "сырых" и "редактируемых" данных
        - формул
        - стилей
        - метаданных диаграмм

        Args:
            sheet_name: Имя листа.

        Returns:
            SheetData: Подготовленные данные для листа.
        """
        # TODO: Реализовать логику извлечения данных из self.db_storage
        # Это заглушка для демонстрации структуры.

        # Пример данных
        sample_data = [
            ["Product", "Q1", "Q2", "Q3"],
            ["Apples", "100", "120", "150"],
            ["Oranges", "80", "90", "100"],
        ]

        sample_formulas = [
            Formula(cell="E2", formula="=SUM(B2:D2)"),
            Formula(cell="E3", formula="=SUM(B3:D3)"),
        ]

        sample_charts = [
            Chart(
                type="col",
                position="A6",
                title="Sales by Quarter",
                series=[
                    ChartSeries(
                        name="Apples",
                        categories="Sheet1!$B$1:$D$1",
                        values="Sheet1!$B$2:$D$2"
                    ),
                    ChartSeries(
                        name="Oranges",
                        categories="Sheet1!$B$1:$D$1",
                        values="Sheet1!$B$3:$D$3"
                    )
                ]
            )
        ]

        return SheetData(
            name=sheet_name,
            data=sample_data,
            formulas=sample_formulas,
            charts=sample_charts
        )

    def _prepare_export_data(self) -> ExportData:
        """
        Подготавливает полную структуру данных для экспорта.

        Returns:
            ExportData: Валидированные данные для передачи в Go.
        """
        # TODO: Получить список всех листов из БД
        sheet_names = ["Sheet1"]  # Заглушка

        sheets_data = []
        for sheet_name in sheet_names:
            sheet_data = self._prepare_sheet_data(sheet_name)
            sheets_data.append(sheet_data)

        # TODO: Получить метаданные проекта из БД
        metadata = ProjectMetadata(
            project_name="Sample Project",
            author="User",
            created_at="2025-09-27T10:00:00Z"
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