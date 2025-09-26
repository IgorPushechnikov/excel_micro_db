# src/core/app_controller.py
"""
Модуль для центрального контроллера приложения.
Управляет жизненным циклом приложения и координирует работу
между различными менеджерами (DataManager, AnalysisManager и т.д.).
"""

import os
import logging
import sqlite3
from typing import Dict, Any, List, Optional, Tuple, Union
from pathlib import Path

# Импортируем анализатор
# from src.analyzer.logic_documentation import analyze_excel_file # Импорт будет в AnalysisManager

# Импортируем хранилище
from src.storage.base import ProjectDBStorage

# Импортируем экспортёры
# from src.exporter.excel.xlsxwriter_exporter import export_project_xlsxwriter as export_with_xlsxwriter # Импорт будет в ExportManager

# Импортируем logger из utils
from src.utils.logger import get_logger

# --- Исключения ---
from src.exceptions.app_exceptions import ProjectError, AnalysisError, ExportError

# Импорты для новых менеджеров (теперь из поддиректории)
# from .controller.data_manager import DataManager # <-- УДАЛЯЕМ ОТСЮДА, чтобы избежать циклического импорта
from .project_manager import ProjectManager # Был перемещен в корень core
# from .controller.format_manager import FormatManager # Пока не реализован
# from .controller.chart_manager import ChartManager # Пока не реализован
# from .controller.analysis_manager import AnalysisManager # Пока не реализован
# from .controller.export_manager import ExportManager # Пока не реализован
# from .controller.node_manager import NodeManager # Пока не реализован

logger = get_logger(__name__)

class AppController:
    """
    Центральный контроллер приложения.

    Координирует работу анализатора, хранилища, процессора и экспортера.
    """

    def __init__(self, project_path: str):
        """
        Инициализирует контроллер для проекта.

        Args:
            project_path (str): Путь к директории проекта.
        """
        self.project_path = project_path
        self.project_db_path = os.path.join(project_path, "project_data.db")
        self.storage: Optional[ProjectDBStorage] = None
        self._current_project_data: Optional[Dict[str, Any]] = None  # Кэш метаданных проекта

        # --- Инициализация менеджеров ---
        # Импортируем DataManager локально, чтобы избежать циклического импорта
        from .controller.data_manager import DataManager # <-- ДОБАВЛЯЕМ СЮДА
        self.data_manager = DataManager(self)

        self.project_manager = ProjectManager(self)
        # self.format_manager = FormatManager(self) # Пока не реализован
        # self.chart_manager = ChartManager(self) # Пока не реализован
        # self.analysis_manager = AnalysisManager(self) # Пока не реализован
        # self.export_manager = ExportManager(self) # Пока не реализован
        # self.node_manager = NodeManager(self) # Пока не реализован

        logger.debug(f"AppController инициализирован для проекта: {project_path}")

    def initialize(self) -> bool:
        """
        Инициализирует контроллер приложения.
        
        Если project_path указан и существует, пытается загрузить проект.
        Если project_path не указан или проект не существует, готов к созданию нового проекта.
        
        Returns:
            bool: True, если инициализация прошла успешно.
        """
        return self.project_manager.initialize()

    @property
    def is_project_loaded(self) -> bool:
        """Проверяет, загружен ли проект."""
        return self.storage is not None

    @property
    def current_project(self) -> Optional[Dict[str, Any]]:
        """Возвращает текущие метаданные проекта."""
        return self._current_project_data
