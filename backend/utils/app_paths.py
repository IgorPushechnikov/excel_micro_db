# backend/utils/app_paths.py
"""
Модуль для определения путей к системным каталогам приложения, таким как AppData.
Используется для хранения глобальных ресурсов (шаблонов, настроек, библиотек).
"""
import platform
from pathlib import Path
import logging
# Импортируем os для доступа к переменным окружения (например, XDG_DATA_HOME на Linux)
import os

# Получаем логгер для этого модуля
from backend.utils.logger import get_logger

logger = get_logger(__name__)


def get_app_data_directory(app_name: str = "ExcelMicroDB") -> Path:
    """
    Определяет стандартный путь к каталогу данных приложения в зависимости от ОС.

    Args:
        app_name (str): Имя приложения, используется для создания подкаталога.
                        По умолчанию "ExcelMicroDB".

    Returns:
        Path: Объект Path, представляющий путь к каталогу данных приложения.
              Каталог будет создан, если он не существует.
    """
    system = platform.system()
    app_dir: Path

    if system == "Windows":
        # На Windows обычно используем %APPDATA% или %LOCALAPPDATA%
        # APPDATA = Roaming, LOCALAPPDATA = Local
        # Для данных приложения, которые не должны синхронизироваться,
        # лучше подходит Local. Для настроек, которые должны следовать за
        # пользователем, Roaming более подходящий.
        # Здесь выберем Roaming для большей переносимости настроек/шаблонов.
        appdata_path = Path.home() / "AppData" / "Roaming"
        app_dir = appdata_path / app_name

    elif system == "Linux":
        # На Linux следуем стандарту XDG Base Directory Specification
        # XDG_DATA_HOME для пользовательских данных
        xdg_data_home = os.environ.get("XDG_DATA_HOME")
        if xdg_data_home:
            base_path = Path(xdg_data_home)
        else:
            # По умолчанию ~/.local/share
            base_path = Path.home() / ".local" / "share"
        app_dir = base_path / app_name

    elif system == "Darwin":  # macOS
        # На macOS используем ~/Library/Application Support
        app_dir = Path.home() / "Library" / "Application Support" / app_name

    else:
        # Для других систем используем fallback - скрытый каталог в домашней директории
        logger.warning(f"Неизвестная ОС ({system}), используем fallback путь.")
        app_dir = Path.home() / f".{app_name.lower()}"

    # Создаем каталог, если он не существует
    try:
        app_dir.mkdir(parents=True, exist_ok=True)
        logger.debug(f"Каталог AppData определен и готов: {app_dir}")
    except Exception as e:
        logger.error(f"Не удалось создать каталог AppData '{app_dir}': {e}")
        # В случае ошибки создания, возвращаем путь, но это может привести к ошибкам далее
        # Пользовательский код должен обрабатывать это
        
    return app_dir

# Пример использования (можно удалить или закомментировать перед интеграцией)
# if __name__ == "__main__":
#     # Для тестирования внутри модуля
#     test_dir = get_app_data_directory("ExcelMicroDB_Test")
#     print(f"Тестовый каталог AppData: {test_dir}")
#     print(f"Существует: {test_dir.exists()}")