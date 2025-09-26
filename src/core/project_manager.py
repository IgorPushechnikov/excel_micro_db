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

    def __init__(self, app_controller):
        """Инициализация менеджера проектов.
        
        Args:
            app_controller: Ссылка на AppController.
        """
        self.app_controller = app_controller
        logger.debug("Инициализация ProjectManager с AppController")

    def initialize(self) -> bool:
        """
        Инициализирует ProjectManager.
        Проверяет, существует ли проект в app_controller.project_path.
        Если существует и валиден, загружает его.
        
        Returns:
            bool: True, если инициализация прошла успешно (проект загружен или готов к созданию).
        """
        project_path = Path(self.app_controller.project_path)
        logger.info(f"Инициализация ProjectManager для проекта: {project_path}")
        
        # Проверяем, существует ли директория проекта
        if not project_path.exists():
            logger.info(f"Директория проекта не существует: {project_path}. Готов к созданию нового проекта.")
            # Не загружаем, но инициализация считается успешной
            return True

        # Проверяем, является ли директория проектом (валидируем структуру и БД)
        if self.validate_project():
            logger.info(f"Проект валиден. Попытка загрузки...")
            if self.load_project():
                logger.info("Проект успешно загружен.")
                return True
            else:
                logger.error("Не удалось загрузить валидный проект.")
                return False
        else:
            logger.warning(f"Директория {project_path} существует, но не является валидным проектом Excel Micro DB.")
            # Можно предложить создать новый проект поверх или вернуть False
            # Пока вернем True, чтобы AppController мог создать проект в этой директории
            return True

    def create_project(self, project_path: str, project_name: Optional[str] = None) -> bool:
        """
        Создание нового проекта.

        Args:
            project_path (str): Путь к директории проекта
            project_name (Optional[str]): Название проекта (по умолчанию имя директории)

        Returns:
            bool: True если проект создан успешно, False в противном случае
        """

        # Инициализируем переменные заранее, чтобы избежать ошибок Pylance
        db_path = None
        storage = None
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

            # --- НОВЫЙ ШАГ: Инициализация БД проекта ---
            logger.debug("Инициализация внутренней БД проекта...")
            try:
                from src.storage.base import ProjectDBStorage
                db_path = project_path_obj / "project_data.db"
                storage = ProjectDBStorage(str(db_path))
                # Главное изменение: вызываем initialize_project_tables, а не пытаемся подключиться
                if not storage.initialize_project_tables():
                    logger.error(f"Не удалось инициализировать схему БД проекта: {db_path}")
                    # Опционально: удалить частично созданные файлы/папки
                    # self._cleanup_partial_project(project_path_obj)
                    return False

                # Отключаемся после инициализации
                storage.disconnect()
                logger.info(f"БД проекта инициализирована: {db_path}")
            except Exception as e:
                logger.error(f"Ошибка при инициализации БД проекта {db_path}: {e}", exc_info=True)
                return False

            # --- КОНЕЦ НОВОГО ШАГА ---

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

    def load_project(self) -> bool:
        """
        Загрузка существующего проекта из app_controller.project_path.
        Инициализирует app_controller.storage.

        Returns:
            bool: True если проект загружен успешно, False в противном случае
        """
        from src.storage.base import ProjectDBStorage

        try:
            project_path_obj = Path(self.app_controller.project_path).resolve()
            logger.info(f"Загрузка проекта из: {project_path_obj}")

            # Проверяем существование проекта
            if not project_path_obj.exists():
                logger.error(f"Директория проекта не существует: {project_path_obj}")
                return False

            if not project_path_obj.is_dir():
                logger.error(f"Путь {project_path_obj} не является директорией")
                return False

            # Проверяем наличие файла метаданных
            metadata_file = project_path_obj / self.PROJECT_METADATA_FILE
            if not metadata_file.exists():
                logger.error(f"Файл метаданных проекта не найден: {metadata_file}")
                logger.info("Возможно, это не проект Excel Micro DB")
                return False

            # Загружаем метаданные
            metadata = self._load_project_metadata(project_path_obj)
            if not metadata:
                return False

            # Обновляем время последнего открытия
            metadata["last_opened"] = datetime.now().isoformat()
            self._save_project_metadata(project_path_obj, metadata)

            # --- НОВОЕ: Инициализация storage для AppController ---
            db_path = project_path_obj / "project_data.db"
            if not db_path.exists():
                logger.error(f"Файл БД проекта не найден: {db_path}")
                return False

            storage = ProjectDBStorage(str(db_path))
            if not storage.connect():
                logger.error(f"Не удалось подключиться к БД проекта: {db_path}")
                return False

            # Устанавливаем storage в app_controller
            self.app_controller.storage = storage
            # Кэшируем метаданные проекта
            self.app_controller._current_project_data = metadata

            logger.info(f"Проект '{metadata.get('project_name')}' загружен успешно, БД подключена")
            return True
            # --- КОНЕЦ НОВОГО ---

        except Exception as e:
            logger.error(f"Ошибка при загрузке проекта: {e}")
            return False

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

    def validate_project(self) -> bool:
        """
        Валидация структуры проекта в app_controller.project_path.

        Returns:
            bool: True если проект валиден, False в противном случае
        """

        # Инициализируем переменные заранее, чтобы избежать ошибок Pylance
        db_path = None
        storage = None
        cursor = None
        try:
            project_path_obj = Path(self.app_controller.project_path)
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

            # --- НОВАЯ ПРОВЕРКА: Валидация БД проекта ---
            logger.debug("Валидация внутренней БД проекта...")
            try:
                db_path = project_path_obj / "project_data.db"
                if not db_path.exists():
                    logger.error(f"Файл БД проекта не найден: {db_path}")
                    return False

                # Пробуем подключиться и выполнить базовую проверку
                from src.storage.base import ProjectDBStorage
                storage = ProjectDBStorage(str(db_path))
                if not storage.connect():
                    logger.error(f"Не удалось подключиться к БД проекта для валидации: {db_path}")
                    return False

                # Проверка наличия ключевых таблиц
                # Убедимся, что соединение установлено и объект cursor можно получить
                if storage.connection:
                    cursor = storage.connection.cursor()
                else:
                    logger.error(f"Не удалось получить соединение с БД проекта: {db_path}")
                    storage.disconnect()
                    return False

                required_tables = ['projects', 'sheets', 'formulas', 'sheet_styles', 'sheet_charts', 'edit_history']
                if cursor:  # Дополнительная проверка на None для Pylance
                    for table_name in required_tables:
                        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (table_name,))
                        if not cursor.fetchone():
                            logger.error(f"БД проекта {db_path} не содержит обязательную таблицу '{table_name}'.")
                            storage.disconnect()
                            return False
                else:
                    logger.error("Не удалось получить курсор для БД проекта")
                    storage.disconnect()
                    return False

                storage.disconnect()
                logger.debug(f"БД проекта {db_path} прошла базовую валидацию.")
            except Exception as e:
                logger.error(f"Ошибка при валидации БД проекта {db_path}: {e}", exc_info=True)
                # Попробуем отключиться на случай ошибки
                try:
                    if storage:
                        storage.disconnect()
                except:
                    pass
                return False

            # --- КОНЕЦ НОВОЙ ПРОВЕРКИ ---

            logger.debug(f"Проект {project_path_obj} прошел валидацию")
            return True
        except Exception as e:
            logger.error(f"Ошибка при валидации проекта: {e}")
            return False

    def close_project(self) -> None:
        """
        Закрытие текущего проекта.
        Отключает app_controller.storage.
        """
        logger.info("Закрытие проекта...")
        if self.app_controller.storage:
            try:
                self.app_controller.storage.disconnect()
                logger.debug("Соединение с БД проекта закрыто.")
            except Exception as e:
                logger.error(f"Ошибка при закрытии соединения с БД проекта: {e}")
            finally:
                self.app_controller.storage = None
                self.app_controller._current_project_data = None
        else:
            logger.debug("Соединение с БД проекта не было установлено.")

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


class _FakeAppController:
    """Фиктивный AppController для standalone запуска ProjectManager."""
    def __init__(self, project_path=""):
        self.project_path = project_path
        # storage и другие атрибуты могут быть добавлены при необходимости


def create_project_manager(app_controller=None) -> ProjectManager:
    """
    Фабричная функция для создания экземпляра ProjectManager.

    Args:
        app_controller: Экземпляр AppController. Если None, создаётся фиктивный.

    Returns:
        ProjectManager: Экземпляр менеджера проектов
    """
    if app_controller is None:
        app_controller = _FakeAppController()
    return ProjectManager(app_controller)


# Пример использования
if __name__ == "__main__":
    # Это просто для демонстрации, не будет выполняться при импорте
    logger.info("Демонстрация работы ProjectManager")

    # Создание менеджера проектов (с фиктивным app_controller)
    test_project_path = "./test_project"
    pm = create_project_manager(_FakeAppController(project_path=test_project_path))

    # Создание тестового проекта (project_path уже в app_controller)
    if pm.create_project(test_project_path, "Тестовый проект"):
        logger.info("Тестовый проект создан успешно")

        # Загрузка проекта (project_path берётся из app_controller)
        if pm.load_project(): # Вызов без аргумента
            logger.info(f"Загружен проект: {pm.app_controller._current_project_data.get('project_name')}")

            # Валидация проекта (project_path берётся из app_controller)
            if pm.validate_project(): # Вызов без аргумента
                logger.info("Проект прошёл валидацию")
            else:
                logger.error("Ошибка при валидации проекта")
        else:
            logger.error("Ошибка при загрузке проекта")