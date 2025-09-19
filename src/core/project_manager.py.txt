# src/core/project_manager.py
"""
Менеджер проектов для Excel Micro DB.
Отвечает за создание, загрузку, сохранение и управление проектами.
"""

import os
import sys
import json
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger

# Получаем логгер для этого модуля
logger = get_logger(__name__)


class ProjectManager:
    """
    Менеджер проектов Excel Micro DB.
    
    Отвечает за:
    - Создание новых проектов
    - Загрузку существующих проектов
    - Сохранение состояния проектов
    - Управление метаданными проектов
    - Валидацию структуры проектов
    """
    
    # Имя файла с метаданными проекта
    PROJECT_METADATA_FILE = ".excel_micro_db_project"
    
    def __init__(self):
        """Инициализация менеджера проектов."""
        logger.debug("Инициализация ProjectManager")
        
    def create_project(self, project_path: str, project_name: Optional[str] = None) -> bool:
        """
        Создание нового проекта.
        
        Args:
            project_path (str): Путь к директории проекта
            project_name (Optional[str]): Название проекта (по умолчанию имя директории)
            
        Returns:
            bool: True если проект создан успешно, False в противном случае
        """
        try:
            project_path_obj = Path(project_path).resolve()
            logger.info(f"Создание нового проекта в: {project_path_obj}")
            
            # Проверяем, что директория существует или создаем её
            if project_path_obj.exists():
                if not project_path_obj.is_dir():
                    logger.error(f"Путь {project_path_obj} существует, но это не директория")
                    return False
                logger.warning(f"Директория {project_path_obj} уже существует")
            else:
                project_path_obj.mkdir(parents=True, exist_ok=True)
                logger.debug(f"Создана директория проекта: {project_path_obj}")
            
            # Если имя проекта не задано, используем имя директории
            if not project_name:
                project_name = project_path_obj.name
            
            # Создаем стандартную структуру проекта
            self._create_project_structure(project_path_obj)
            
            # Создаем метаданные проекта
            metadata = self._create_project_metadata(project_name, str(project_path_obj))
            if not self._save_project_metadata(project_path_obj, metadata):
                logger.error("Не удалось сохранить метаданные проекта")
                return False
            
            logger.info(f"Проект '{project_name}' создан успешно в: {project_path_obj}")
            return True
            
        except PermissionError:
            logger.error(f"Нет прав для создания проекта в: {project_path}")
            return False
        except Exception as e:
            logger.error(f"Ошибка при создании проекта: {e}")
            return False
    
    def _create_project_structure(self, project_path: Path) -> None:
        """
        Создание стандартной структуры директорий проекта.
        
        Args:
            project_path (Path): Путь к директории проекта
        """
        # Определяем необходимые поддиректории
        subdirs = [
            "data/raw",
            "data/processed", 
            "data/samples",
            "output/documentation",
            "output/reports",
            "output/exports",
            "config",
            "logs"
        ]
        
        # Создаем каждую поддиректорию
        for subdir in subdirs:
            (project_path / subdir).mkdir(parents=True, exist_ok=True)
            logger.debug(f"Создана директория: {project_path / subdir}")
        
        logger.debug("Структура проекта создана")
    
    def _create_project_metadata(self, project_name: str, project_path: str) -> Dict[str, Any]:
        """
        Создание метаданных проекта.
        
        Args:
            project_name (str): Название проекта
            project_path (str): Путь к проекту
            
        Returns:
            Dict[str, Any]: Словарь с метаданными проекта
        """
        now = datetime.now().isoformat()
        metadata = {
            "project_name": project_name,
            "project_path": project_path,
            "created_at": now,
            "last_opened": now,
            "version": "1.0",
            "excel_files": [],
            "analyses": [],
            "config": {
                "max_file_size_mb": 100,
                "analysis_timeout_seconds": 300
            }
        }
        logger.debug(f"Созданы метаданные проекта для '{project_name}'")
        return metadata
    
    def _save_project_metadata(self, project_path: Path, metadata: Dict[str, Any]) -> bool:
        """
        Сохранение метаданных проекта в файл.
        
        Args:
            project_path (Path): Путь к директории проекта
            metadata (Dict[str, Any]): Метаданные проекта
            
        Returns:
            bool: True если сохранение успешно, False в противном случае
        """
        try:
            metadata_file = project_path / self.PROJECT_METADATA_FILE
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            logger.debug(f"Метаданные проекта сохранены в: {metadata_file}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при сохранении метаданных проекта: {e}")
            return False
    
    def load_project(self, project_path: str) -> Optional[Dict[str, Any]]:
        """
        Загрузка существующего проекта.
        
        Args:
            project_path (str): Путь к директории проекта
            
        Returns:
            Optional[Dict[str, Any]]: Метаданные проекта или None если ошибка
        """
        try:
            project_path_obj = Path(project_path).resolve()
            logger.info(f"Загрузка проекта из: {project_path_obj}")
            
            # Проверяем существование проекта
            if not project_path_obj.exists():
                logger.error(f"Директория проекта не существует: {project_path_obj}")
                return None
            
            if not project_path_obj.is_dir():
                logger.error(f"Путь {project_path_obj} не является директорией")
                return None
            
            # Проверяем наличие файла метаданных
            metadata_file = project_path_obj / self.PROJECT_METADATA_FILE
            if not metadata_file.exists():
                logger.error(f"Файл метаданных проекта не найден: {metadata_file}")
                logger.info("Возможно, это не проект Excel Micro DB")
                return None
            
            # Загружаем метаданные
            metadata = self._load_project_metadata(project_path_obj)
            if not metadata:
                return None
            
            # Обновляем время последнего открытия
            metadata["last_opened"] = datetime.now().isoformat()
            self._save_project_metadata(project_path_obj, metadata)
            
            logger.info(f"Проект '{metadata.get('project_name')}' загружен успешно")
            return metadata
            
        except Exception as e:
            logger.error(f"Ошибка при загрузке проекта: {e}")
            return None
    
    def _load_project_metadata(self, project_path: Path) -> Optional[Dict[str, Any]]:
        """
        Загрузка метаданных проекта из файла.
        
        Args:
            project_path (Path): Путь к директории проекта
            
        Returns:
            Optional[Dict[str, Any]]: Метаданные проекта или None если ошибка
        """
        try:
            metadata_file = project_path / self.PROJECT_METADATA_FILE
            with open(metadata_file, 'r', encoding='utf-8') as f:
                metadata = json.load(f)
            logger.debug(f"Метаданные проекта загружены из: {metadata_file}")
            return metadata
        except json.JSONDecodeError as e:
            logger.error(f"Ошибка формата JSON в файле метаданных: {e}")
            return None
        except Exception as e:
            logger.error(f"Ошибка при загрузке метаданных проекта: {e}")
            return None
    
    def validate_project(self, project_path: str) -> bool:
        """
        Валидация структуры проекта.
        
        Args:
            project_path (str): Путь к директории проекта
            
        Returns:
            bool: True если проект валиден, False в противном случае
        """
        try:
            project_path_obj = Path(project_path)
            logger.debug(f"Валидация проекта: {project_path_obj}")
            
            # Проверяем существование директории
            if not project_path_obj.exists() or not project_path_obj.is_dir():
                logger.error(f"Директория проекта не существует или не является директорией: {project_path_obj}")
                return False
            
            # Проверяем наличие файла метаданных
            metadata_file = project_path_obj / self.PROJECT_METADATA_FILE
            if not metadata_file.exists():
                logger.error(f"Файл метаданных проекта не найден: {metadata_file}")
                return False
            
            # Проверяем наличие обязательных поддиректорий
            required_dirs = ["data", "output", "config"]
            for dir_name in required_dirs:
                if not (project_path_obj / dir_name).exists():
                    logger.warning(f"Отсутствует обязательная директория: {dir_name}")
                    # Не возвращаем False, так как можем попытаться восстановить структуру
            
            logger.debug(f"Проект {project_path_obj} прошел валидацию")
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при валидации проекта: {e}")
            return False
    
    def add_excel_file_to_project(self, project_path: str, file_path: str) -> bool:
        """
        Добавление Excel файла в проект.
        
        Args:
            project_path (str): Путь к директории проекта
            file_path (str): Путь к Excel файлу
            
        Returns:
            bool: True если файл добавлен успешно, False в противном случае
        """
        try:
            project_path_obj = Path(project_path)
            file_path_obj = Path(file_path)
            
            # Загружаем метаданные проекта
            metadata = self._load_project_metadata(project_path_obj)
            if not metadata:
                return False
            
            # Проверяем, что файл существует
            if not file_path_obj.exists():
                logger.error(f"Excel файл не существует: {file_path_obj}")
                return False
            
            # Добавляем файл в список, если его там еще нет
            file_record = {
                "name": file_path_obj.name,
                "path": str(file_path_obj.relative_to(project_path_obj)) if file_path_obj.is_relative_to(project_path_obj) else str(file_path_obj),
                "added_at": datetime.now().isoformat(),
                "size_bytes": file_path_obj.stat().st_size
            }
            
            if file_record not in metadata.get("excel_files", []):
                if "excel_files" not in metadata:
                    metadata["excel_files"] = []
                metadata["excel_files"].append(file_record)
                logger.debug(f"Excel файл добавлен в проект: {file_path_obj}")
            
            # Сохраняем обновленные метаданные
            return self._save_project_metadata(project_path_obj, metadata)
            
        except Exception as e:
            logger.error(f"Ошибка при добавлении Excel файла в проект: {e}")
            return False
    
    def get_project_list(self, root_path: str = ".") -> Dict[str, Dict[str, Any]]:
        """
        Получение списка проектов в указанной директории.
        
        Args:
            root_path (str): Корневая директория для поиска проектов (по умолчанию текущая)
            
        Returns:
            Dict[str, Dict[str, Any]]: Словарь с информацией о проектах
        """
        projects = {}
        root_path_obj = Path(root_path)
        
        try:
            # Рекурсивно ищем проекты в поддиректориях
            for metadata_file in root_path_obj.rglob(self.PROJECT_METADATA_FILE):
                try:
                    project_path = metadata_file.parent
                    with open(metadata_file, 'r', encoding='utf-8') as f:
                        metadata = json.load(f)
                    
                    project_name = metadata.get("project_name", project_path.name)
                    projects[str(project_path)] = {
                        "name": project_name,
                        "path": str(project_path),
                        "created_at": metadata.get("created_at"),
                        "last_opened": metadata.get("last_opened"),
                        "excel_files_count": len(metadata.get("excel_files", []))
                    }
                except Exception as e:
                    logger.warning(f"Ошибка при чтении проекта {metadata_file.parent}: {e}")
                    continue
            
            logger.debug(f"Найдено {len(projects)} проектов в {root_path}")
            return projects
            
        except Exception as e:
            logger.error(f"Ошибка при поиске проектов: {e}")
            return projects
    
    def cleanup(self) -> None:
        """
        Очистка ресурсов менеджера проектов.
        """
        logger.debug("Очистка ресурсов ProjectManager")
        # Здесь можно добавить закрытие файлов, соединений и т.д.


# Функции для удобного использования менеджера проектов
def create_project_manager() -> ProjectManager:
    """
    Фабричная функция для создания экземпляра ProjectManager.
    
    Returns:
        ProjectManager: Экземпляр менеджера проектов
    """
    return ProjectManager()


# Пример использования
if __name__ == "__main__":
    # Это просто для демонстрации, не будет выполняться при импорте
    logger.info("Демонстрация работы ProjectManager")
    
    # Создание менеджера проектов
    pm = create_project_manager()
    
    # Создание тестового проекта
    test_project_path = "./test_project"
    if pm.create_project(test_project_path, "Тестовый проект"):
        logger.info("Тестовый проект создан успешно")
        
        # Загрузка проекта
        project_data = pm.load_project(test_project_path)
        if project_data:
            logger.info(f"Загружен проект: {project_data.get('project_name')}")
        
        # Валидация проекта
        if pm.validate_project(test_project_path):
            logger.info("Проект прошел валидацию")
    else:
        logger.error("Ошибка при создании тестового проекта")