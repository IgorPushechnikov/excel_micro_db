# scripts/run_integration_test.py
"""
Скрипт для автоматического запуска интеграционного теста
(инициализация -> анализ -> экспорт) с сохранением лога в файл.
"""

import subprocess
import sys
from pathlib import Path
import logging
from datetime import datetime

def setup_logging(log_file_path: Path):
    """Настраивает логирование для скрипта."""
    # Настройка логгирования Python
    logging.basicConfig(
        level=logging.DEBUG,  # Уровень логирования DEBUG для макс. детализации
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file_path, encoding='utf-8'), # В файл
            logging.StreamHandler(sys.stdout) # И в консоль
        ],
        force=True # Перезаписать любую существующую конфигурацию
    )
    return logging.getLogger(__name__)

def run_command(command: list, description: str, logger) -> bool:
    """
    Запускает команду и логирует её выполнение.
    Args:
        command (list): Команда в виде списка строк.
        description (str): Описание команды для лога.
        logger: Логгер для записи сообщений.
    Returns:
        bool: True, если команда выполнена успешно, False в противном случае.
    """
    logger.info(f"--- Начало: {description} ---")
    logger.info(f"Команда: {' '.join(command)}")
    
    try:
        # Используем subprocess.run для выполнения команды
        # capture_output=True захватывает stdout и stderr
        # text=True возвращает строки, а не байты
        # timeout=120 даст команде 120 секунд на выполнение
        # cwd=project_root для запуска из корня проекта
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=120,
            cwd=project_root 
        )
        
        # Логируем stdout и stderr
        if result.stdout:
            logger.info(f"STDOUT:\n{result.stdout}")
        if result.stderr:
            # Логируем stderr как WARNING или ERROR, в зависимости от кода возврата
            if result.returncode == 0:
                logger.warning(f"STDERR (код 0, но есть вывод):\n{result.stderr}")
            else:
                logger.error(f"STDERR:\n{result.stderr}")
                
        if result.returncode == 0:
            logger.info(f"--- Успех: {description} ---")
            return True
        else:
            logger.error(f"--- Ошибка: {description} (Код возврата: {result.returncode}) ---")
            return False
            
    except subprocess.TimeoutExpired:
        logger.error(f"--- Ошибка: {description} превысила тайм-аут (120 секунд) ---")
        return False
    except Exception as e:
        logger.error(f"--- Ошибка: Не удалось выполнить {description}: {e} ---")
        return False

def main():
    """Главная функция скрипта."""
    # --- Настройка логгирования в файл ---
    # Определяем корневую директорию проекта
    global project_root
    project_root = Path(__file__).parent.parent.resolve() # resolve() для получения абсолютного пути

    # Создаем директорию для логов тестов, если её нет, внутри test_workspace проекта
    log_dir = project_root / "test_workspace" / "integration_test_logs"
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        # print(f"[СКРИПТ] Директория для логов: {log_dir}") # Выводим до настройки логгера
    except Exception as e:
        print(f"[СКРИПТ] Ошибка создания директории для логов {log_dir}: {e}")
        sys.exit(1)

    # Создаем имя файла лога с текущей датой и временем
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"integration_test_{timestamp}.log"
    log_file_path = log_dir / log_file_name

    # Настройка логгирования
    logger = setup_logging(log_file_path)
    # --- Конец настройки логгирования ---

    logger.info("=== НАЧАЛО ИНТЕГРАЦИОННОГО ТЕСТА ===")
    logger.info(f"Корневая директория проекта определена как: {project_root}")
    logger.info(f"Лог скрипта будет сохранен в: {log_file_path}")
    print(f"[СКРИПТ] Лог скрипта будет сохранен в: {log_file_path}") # Выводим путь сразу
    
    # Определяем пути относительно корня проекта
    project_name = f"integration_test_project_{timestamp}"
    project_path = project_root / "test_workspace" / project_name
    excel_file = project_root / "data" / "samples" / "test_sample.xlsx"
    output_file = project_path / "output" / "exports" / "recreated_test_sample.xlsx"
    main_py = project_root / "main.py"

    # Проверяем существование необходимых файлов
    if not main_py.exists():
        logger.critical(f"Файл main.py не найден: {main_py}")
        print(f"[СКРИПТ] Критическая ошибка: Файл main.py не найден: {main_py}")
        sys.exit(1)
    if not excel_file.exists():
        logger.critical(f"Тестовый файл Excel не найден: {excel_file}")
        print(f"[СКРИПТ] Критическая ошибка: Тестовый файл Excel не найден: {excel_file}")
        sys.exit(1)

    logger.info(f"Проект будет создан в: {project_path}")
    logger.info(f"Тестовый файл: {excel_file}")
    logger.info(f"Выходной файл: {output_file}")

    all_steps_passed = True

    # 1. Инициализация проекта
    cmd_init = [
        sys.executable, str(main_py),
        "--init",
        "--project-path", str(project_path)
    ]
    if not run_command(cmd_init, "Инициализация проекта", logger):
        all_steps_passed = False

    # 2. Анализ Excel-файла (только если инициализация успешна)
    if all_steps_passed:
        cmd_analyze = [
            sys.executable, str(main_py),
            "--analyze", str(excel_file),
            "--project-path", str(project_path)
        ]
        if not run_command(cmd_analyze, "Анализ Excel-файла", logger):
            all_steps_passed = False

    # 3. Экспорт результатов (только если анализ успешен)
    if all_steps_passed:
        cmd_export = [
            sys.executable, str(main_py),
            "--export", "excel",
            "--output", str(output_file),
            "--project-path", str(project_path)
        ]
        if not run_command(cmd_export, "Экспорт результатов", logger):
            all_steps_passed = False

    # --- Проверка результатов ---
    if all_steps_passed:
        logger.info("=== ВСЕ ЭТАПЫ ТЕСТА ЗАВЕРШЕНЫ ===")
        db_file = project_path / "project_data.db"
        if db_file.exists():
            logger.info(f"✓ Файл БД проекта создан: {db_file}")
        else:
            logger.error(f"✗ Файл БД проекта НЕ НАЙДЕН: {db_file}")
            all_steps_passed = False

        if output_file.exists():
            logger.info(f"✓ Выходной Excel-файл создан: {output_file}")
        else:
            logger.error(f"✗ Выходной Excel-файл НЕ НАЙДЕН: {output_file}")
            all_steps_passed = False
        
        # --- Копирование лога теста в папку экспорта ---
        if all_steps_passed:
            try:
                import shutil
                # Определяем путь к папке экспорта
                output_exports_dir_path = output_file.parent
                # Определяем путь к папке логов внутри экспорта
                export_logs_dir = output_exports_dir_path / "logs"
                export_logs_dir.mkdir(parents=True, exist_ok=True) # Создаем папку logs, если её нет
                
                # Определяем новое имя файла лога (такое же, как оригинальный)
                copied_log_file_path = export_logs_dir / log_file_path.name
                
                # Копируем файл лога
                shutil.copy2(log_file_path, copied_log_file_path)
                logger.info(f"✓ Лог интеграционного теста скопирован в папку экспорта: {copied_log_file_path}")
                print(f"[СКРИПТ] Лог интеграционного теста скопирован в: {copied_log_file_path}")
            except Exception as copy_error:
                logger.error(f"✗ Не удалось скопировать лог интеграционного теста в папку экспорта: {copy_error}")
                print(f"[СКРИПТ] Ошибка при копировании лога теста: {copy_error}")
                # Не считаем это критической ошибкой для всего теста, поэтому all_steps_passed не меняем
        # --- Конец копирования лога ---
    else:
        logger.error("=== ТЕСТ ЗАВЕРШЕН С ОШИБКАМИ ===")

    logger.info(f"Подробный лог сохранен в: {log_file_path}")
    print(f"\n[СКРИПТ] Путь к лог-файлу этого запуска: {log_file_path}") # Всегда выводим путь в консоль
    if all_steps_passed:
        logger.info("Тест пройден успешно!")
        print("[СКРИПТ] Тест пройден успешно!")
        sys.exit(0)
    else:
        logger.error("Тест провален!")
        print("[СКРИПТ] Тест провален!")
        sys.exit(1)

if __name__ == "__main__":
    main()