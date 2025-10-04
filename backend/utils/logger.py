# backend/utils/logger.py
"""
Модуль для настройки и предоставления логгера для всего приложения Excel Micro DB.
"""

import logging
import sys
import os
from pathlib import Path
from typing import Optional

# --- Настройки логгера ---
# Базовый уровень логирования (DEBUG, INFO, WARNING, ERROR, CRITICAL)
# В production его можно изменить, например, на WARNING или ERROR
BASE_LOG_LEVEL = logging.DEBUG

# Формат лога (можно настроить под свои нужды)
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'

# Путь к файлу лога (относительно корня проекта или абсолютный)
# LOG_FILE_PATH = "logs/app.log" # <-- Относительный путь, создастся в корне проекта
# ИЛИ
# LOG_FILE_PATH = "/var/log/excel_micro_db/app.log" # <-- Абсолютный путь (для Linux/Server)
# ИЛИ
# LOG_FILE_PATH = os.path.join(os.path.dirname(__file__), '..', '..', 'logs', 'app.log') # <-- Относительный путь, вычисляемый

# Уровень логирования для файлового хендлера (может отличаться от консольного)
FILE_LOG_LEVEL = logging.DEBUG

# Уровень логирования для консольного хендлера
CONSOLE_LOG_LEVEL = logging.INFO

# --- НОВОЕ: Глобальная переменная для состояния логирования ---
_LOGGING_ENABLED = True # По умолчанию включено
# --- КОНЕЦ НОВОГО ---

# --- Глобальные переменные для логгера ---
# _logger_instance будет хранить экземпляр корневого логгера приложения
_logger_instance: Optional[logging.Logger] = None

# _log_file_path будет хранить путь к файлу лога, если он используется
_log_file_path: Optional[str] = None


def setup_logger(log_file_path: Optional[str] = None, force_recreate: bool = False) -> logging.Logger:
    """
    Настраивает и возвращает корневой логгер для приложения Excel Micro DB.

    Эта функция должна вызываться один раз при запуске приложения (в main.py или gui.py)
    для инициализации логгирования. Последующие вызовы get_logger будут
    возвращать уже настроенный логгер.

    Args:
        log_file_path (Optional[str]): Путь к файлу лога. Если None, логирование в файл отключено.
        force_recreate (bool): Если True, заново создает логгер, даже если он уже существует.
                              Используется в основном для тестов.

    Returns:
        logging.Logger: Настроенный экземпляр корневого логгера приложения.
    """
    global _logger_instance, _log_file_path, _LOGGING_ENABLED # Добавляем _LOGGING_ENABLED в глобальные

    # Проверяем, существует ли уже настроенный логгер
    if _logger_instance is not None and not force_recreate:
        # logger.debug("Логгер уже настроен. Возвращается существующий экземпляр.")
        # Не используем logger.debug здесь, так как _logger_instance еще не инициализирован
        # для этого вызова setup_logger. Просто возвращаем его.
        return _logger_instance

    # --- Создание или получение корневого логгера ---
    # Используем имя "excel_micro_db" для корневого логгера приложения
    # Это позволит легко идентифицировать логи нашего приложения
    logger = logging.getLogger("excel_micro_db")

    # Устанавливаем базовый уровень логирования для логгера
    # Используем глобальную переменную _LOGGING_ENABLED для определения уровня
    logger.setLevel(BASE_LOG_LEVEL if _LOGGING_ENABLED else logging.CRITICAL + 1)

    # --- Очистка существующих хендлеров (если force_recreate=True или при повторной настройке) ---
    # Это важно, чтобы избежать дублирования сообщений в логе при повторных вызовах setup_logger
    if logger.hasHandlers():
        logger.handlers.clear()

    # --- Создание форматтера ---
    formatter = logging.Formatter(LOG_FORMAT)

    # --- Настройка файлового хендлера ---
    if log_file_path and _LOGGING_ENABLED: # Добавляем проверку _LOGGING_ENABLED
        try:
            # Создаем директорию для файла лога, если её нет
            log_file_dir = Path(log_file_path).parent
            log_file_dir.mkdir(parents=True, exist_ok=True)

            # Создаем FileHandler
            # mode='a' означает "append" - добавлять к существующему файлу
            # encoding='utf-8' гарантирует корректную запись кириллицы и других символов
            file_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
            file_handler.setLevel(FILE_LOG_LEVEL if _LOGGING_ENABLED else logging.CRITICAL + 1) # Устанавливаем уровень в зависимости от состояния
            file_handler.setFormatter(formatter)

            # Добавляем FileHandler к логгеру
            logger.addHandler(file_handler)

            # Сохраняем путь к файлу лога для возможного использования
            _log_file_path = log_file_path

            logger.debug(f"FileHandler добавлен. Логи будут записываться в: {log_file_path}")

        except Exception as e:
            # Если не удалось настроить логирование в файл, логируем это в stderr (или в существующий лог, если он уже настроен)
            print(f"Ошибка при настройке FileHandler для лога '{log_file_path}': {e}", file=sys.stderr)
            logger.error(f"Ошибка при настройке FileHandler для лога '{log_file_path}': {e}", exc_info=True)
            # Не прерываем настройку из-за ошибки файла лога

    # --- Настройка консольного хендлера ---
    if _LOGGING_ENABLED: # Добавляем проверку _LOGGING_ENABLED
        console_handler = logging.StreamHandler(sys.stdout) # Используем stdout для INFO и ниже, stderr для WARNING и выше?
        console_handler.setLevel(CONSOLE_LOG_LEVEL if _LOGGING_ENABLED else logging.CRITICAL + 1) # Устанавливаем уровень в зависимости от состояния
        console_handler.setFormatter(formatter)

        # Добавляем ConsoleHandler к логгеру
        logger.addHandler(console_handler)

        logger.debug("ConsoleHandler добавлен.")

    # --- Сохранение экземпляра логгера ---
    _logger_instance = logger

    logger.info("Корневой логгер приложения 'excel_micro_db' успешно настроен.")
    if _log_file_path and _LOGGING_ENABLED:
        logger.info(f"Логирование в файл включено: {_log_file_path}")
    elif _LOGGING_ENABLED:
        logger.info("Логирование в файл отключено.")
    else:
        logger.info("Все логирование отключено.") # Сообщение, если логирование выключено

    return logger


def get_logger(module_name: str) -> logging.Logger:
    """
    Возвращает логгер для конкретного модуля.

    Эта функция используется во всех модулях приложения для получения
    настроенного логгера. Она наследует настройки от корневого логгера,
    созданного setup_logger.

    Args:
        module_name (str): Имя модуля, для которого создается логгер.
                          Обычно передается __name__ из модуля.

    Returns:
        logging.Logger: Логгер для указанного модуля.
    """
    # Возвращаем логгер с именем "excel_micro_db.module_name"
    # Это создаст дочерний логгер, который будет наследовать настройки
    # от корневого логгера "excel_micro_db"
    return logging.getLogger(f"excel_micro_db.{module_name}")


def get_log_file_path() -> Optional[str]:
    """
    Возвращает путь к файлу лога, если он был настроен.

    Returns:
        Optional[str]: Путь к файлу лога или None, если логирование в файл отключено.
    """
    return _log_file_path

# --- НОВОЕ: Функции для включения/отключения логирования ---
def set_logging_enabled(enabled: bool):
    """
    Включает или отключает логирование для всего приложения.

    Args:
        enabled (bool): True для включения, False для отключения.
    """
    global _LOGGING_ENABLED
    _LOGGING_ENABLED = enabled
    logger_instance = logging.getLogger("excel_micro_db")
    if logger_instance and logger_instance.hasHandlers():
        # Устанавливаем уровень для всех хендлеров
        level_to_set = BASE_LOG_LEVEL if enabled else logging.CRITICAL + 1
        for handler in logger_instance.handlers:
            handler.setLevel(level_to_set)
        # Также устанавливаем уровень для самого логгера
        logger_instance.setLevel(level_to_set)
    # Если логгер еще не настроен, изменения вступят в силу при вызове setup_logger
    # Проверим, что _logger_instance существует, прежде чем логировать
    if _logger_instance:
        _logger_instance.info(f"Логирование {'включено' if enabled else 'отключено'}.")

def is_logging_enabled() -> bool:
    """
    Проверяет, включено ли логирование.

    Returns:
        bool: True, если логирование включено, иначе False.
    """
    return _LOGGING_ENABLED
# --- КОНЕЦ НОВОГО ---

# --- Автоматическая настройка логгера при импорте модуля (опционально) ---
# Это может быть полезно для скриптов, которые не вызывают setup_logger напрямую.
# Однако, лучше явно вызывать setup_logger в main.py или gui.py для контроля.
# if __name__ != "__main__":
#     # Проверяем, не настроен ли логгер уже (например, в тестах)
#     if _logger_instance is None:
#         # Пытаемся получить путь к файлу лога из переменной окружения или использовать значение по умолчанию
#         default_log_file = os.environ.get('EXCEL_MICRO_DB_LOG_FILE', 'logs/app.log')
#         setup_logger(default_log_file)

# Пример использования (если файл запускается напрямую)
if __name__ == "__main__":
    # Настройка логгера с логированием в файл
    logger = setup_logger("logs/test_logger.log")
    logger.info("Тестовое сообщение INFO из utils/logger.py")
    logger.debug("Тестовое сообщение DEBUG из utils/logger.py")
    logger.warning("Тестовое сообщение WARNING из utils/logger.py")
    logger.error("Тестовое сообщение ERROR из utils/logger.py")
    # logger.critical("Тестовое сообщение CRITICAL из utils/logger.py")

    # Получение логгера для модуля
    module_logger = get_logger(__name__)
    module_logger.info("Тестовое сообщение INFO от логгера модуля")

    print("Логгер настроен и протестирован.")
