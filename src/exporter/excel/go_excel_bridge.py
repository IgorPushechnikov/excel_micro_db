import subprocess
import logging
import os
from pathlib import Path

# Получаем логгер
logger = logging.getLogger(__name__)


def run_go_excel_exporter(json_input_path: str, xlsx_output_path: str, go_exe_path: str = "go_excel_exporter.exe") -> bool:
    """
    Вызывает Go-бинарник go_excel_exporter.exe для преобразования JSON в XLSX.

    Args:
        json_input_path (str): Путь к входному JSON-файлу.
        xlsx_output_path (str): Путь к выходному XLSX-файлу.
        go_exe_path (str): Путь к исполняемому файлу go_excel_exporter.exe.
                           По умолчанию предполагается, что он находится в PATH
                           или в текущей рабочей директории.
                           Можно указать абсолютный путь.

    Returns:
        bool: True, если экспорт успешен, иначе False.
    """
    logger.info(f"Начало вызова Go-экспортёра: {go_exe_path}")
    logger.debug(f"  Входной JSON: {json_input_path}")
    logger.debug(f"  Выходной XLSX: {xlsx_output_path}")

    # 1. Проверка существования файлов
    if not os.path.exists(json_input_path):
        logger.error(f"Входной JSON-файл не найден: {json_input_path}")
        return False

    # Убедимся, что директория для выходного файла существует
    output_dir = os.path.dirname(xlsx_output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # 2. Подготовка команды
    # Используем список аргументов для лучшей безопасности и обработки путей с пробелами
    cmd = [go_exe_path, "-input", json_input_path, "-output", xlsx_output_path]
    logger.debug(f"Команда для выполнения: {' '.join(cmd)}")

    try:
        # 3. Выполнение команды
        # capture_output=True захватывает stdout/stderr
        # text=True возвращает строки, а не байты
        # timeout=N задаёт максимальное время выполнения (например, 60 секунд)
        # check=True вызовет исключение, если код возврата != 0
        # cwd=... можно указать рабочую директорию, если это важно для бинарника
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120, # 2 минуты на экспорт
            check=True # Вызовет CalledProcessError, если returncode != 0
        )
        
        # 4. Логирование вывода (даже если команда завершилась успешно)
        if result.stdout:
            logger.info(f"STDOUT Go-экспортёра:\n{result.stdout}")
        # STDERR логируем как WARNING или INFO, в зависимости от контекста
        # Поскольку check=True, если дошли до сюда, код 0, stderr может быть не критичным
        if result.stderr:
            logger.info(f"STDERR Go-экспортёра (код 0, но есть вывод):\n{result.stderr}")
            
        logger.info(f"Go-экспортёр успешно завершил работу. Выходной файл: {xlsx_output_path}")
        return True
        
    except subprocess.CalledProcessError as e:
        # Процесс был запущен, но вернул ненулевой код возврата
        logger.error(f"Go-экспортёр завершился с ошибкой (код возврата {e.returncode}).")
        if e.stdout:
            logger.error(f"STDOUT Go-экспортёра:\n{e.stdout}")
        if e.stderr:
            logger.error(f"STDERR Go-экспортёра:\n{e.stderr}")
        return False
        
    except subprocess.TimeoutExpired as e:
        # Процесс не завершился в течение таймаута
        logger.error(f"Go-экспортёр превысил таймаут ({e.timeout} секунд).")
        # e.output и e.stdout могут быть доступны, если процесс успел что-то вывести
        return False
        
    except FileNotFoundError as e:
        # Исполняемый файл не найден
        logger.error(f"Исполняемый файл Go-экспортёра не найден по пути '{go_exe_path}': {e}")
        logger.error("Убедитесь, что go_excel_exporter.exe скомпилирован и находится в указанном пути или в PATH.")
        return False
        
    except Exception as e:
        # Другие ошибки (например, проблемы с правами доступа)
        logger.error(f"Неожиданная ошибка при вызове Go-экспортёра: {e}", exc_info=True)
        return False
