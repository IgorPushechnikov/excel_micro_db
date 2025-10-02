# src/utils/logger.py
"""
Модуль настройки логирования для Excel Micro DB.
Использует конфигурацию из config/settings.yaml.
"""

import logging
import logging.handlers
import os
from pathlib import Path
import yaml

# Путь к конфигурационному файлу
CONFIG_PATH = Path(__file__).parent.parent.parent / "config" / "settings.yaml"

def load_config():
    """Загружает конфигурацию из YAML файла."""
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        print(f"Конфигурационный файл {CONFIG_PATH} не найден. Используются настройки по умолчанию.")
        return {}
    except yaml.YAMLError as e:
        print(f"Ошибка при чтении конфигурационного файла: {e}. Используются настройки по умолчанию.")
        return {}

# Загрузка конфигурации
config = load_config()
logging_config = config.get('logging', {})
app_config = config.get('app', {})

# Получение настроек из конфига или значений по умолчанию
LOG_LEVEL = logging_config.get('level', 'INFO')
LOG_FORMAT = logging_config.get('format', '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
DATE_FORMAT = logging_config.get('date_format', '%Y-%m-%d %H:%M:%S')
LOGS_DIR = Path(config.get('paths', {}).get('logs_dir', './logs'))
CONSOLE_OUTPUT = logging_config.get('console_output', True)
FILE_OUTPUT = logging_config.get('file_output', True)

# Создание директории для логов, если она не существует
LOGS_DIR.mkdir(parents=True, exist_ok=True)

def get_logger(name: str) -> logging.Logger:
    """
    Создает и возвращает настроенный логгер.
    
    Args:
        name (str): Имя логгера (обычно __name__).
        
    Returns:
        logging.Logger: Настроенный экземпляр логгера.
    """
    logger = logging.getLogger(name)
    
    # Установка уровня логирования
    logger.setLevel(getattr(logging, LOG_LEVEL.upper(), logging.DEBUG))
    
    # Предотвращение дублирования обработчиков при повторных вызовах
    if logger.hasHandlers():
        logger.handlers.clear()
    
    # Создание форматтера
    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)
    
    # Консольный обработчик
    if CONSOLE_OUTPUT:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(getattr(logging, LOG_LEVEL.upper(), logging.DEBUG))
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    # Файловый обработчик с ротацией
    if FILE_OUTPUT:
        log_file = LOGS_DIR / f"{app_config.get('name', 'app').lower().replace(' ', '_')}.log"
        # Используем RotatingFileHandler для ротации по размеру
        file_handler = logging.handlers.RotatingFileHandler(
            log_file,
            maxBytes=50*1024*1024,  # 50MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(getattr(logging, LOG_LEVEL.upper(), logging.DEBUG))
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger

# Пример использования:
if __name__ == "__main__":
    # Это просто для демонстрации, не будет выполняться при импорте
    logger = get_logger(__name__)
    logger.debug("Это сообщение уровня DEBUG")
    logger.info("Это сообщение уровня INFO")
    logger.warning("Это сообщение уровня WARNING")
    logger.error("Это сообщение уровня ERROR")
    logger.critical("Это сообщение уровня CRITICAL")