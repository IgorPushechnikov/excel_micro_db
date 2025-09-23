#!/usr/bin/env python3
"""
Точка входа в приложение Excel Micro DB (Командная строка).
"""

import argparse
import sys
import os
import subprocess
from pathlib import Path

# Добавляем src в путь Python, чтобы можно было импортировать модули оттуда
sys.path.insert(0, str(Path(__file__).resolve().parent / 'src'))

from utils.logger import setup_logger # noqa: E402
from core.project_manager import ProjectManager # noqa: E402
from core.app_controller import get_app_controller # noqa: E402
from exceptions.app_exceptions import AppException # noqa: E402

logger = setup_logger(__name__)

def init_project(args):
    """Инициализирует новый проект."""
    logger.info(f"Инициализация проекта в: {args.project_path}")
    project_path = Path(args.project_path).resolve()
    controller = get_app_controller(str(project_path), args.log_level) # Исправленный вызов
    try:
        controller.init_project()
        print(f"Проект успешно инициализирован в {project_path}")
    except Exception as e:
        logger.error(f"Ошибка при инициализации проекта: {e}")
        sys.exit(1)

def analyze_file(args):
    """Анализирует указанный Excel-файл."""
    logger.info(f"Анализ файла: {args.file_path}")
    file_path = Path(args.file_path).resolve()
    project_path = Path(args.project_path).resolve()
    if not file_path.exists():
        logger.error(f"Файл не найден: {file_path}")
        sys.exit(1)
    controller = get_app_controller(str(project_path), args.log_level) # Исправленный вызов
    try:
        controller.analyze_file(str(file_path))
        print(f"Файл {file_path} успешно проанализирован и данные сохранены в проект {project_path}")
    except Exception as e:
        logger.error(f"Ошибка при анализе файла: {e}")
        sys.exit(1)

def export_project(args):
    """Экспортирует проект в Excel."""
    logger.info(f"Экспорт проекта из: {args.project_path}")
    project_path = Path(args.project_path).resolve()
    output_path = Path(args.output_path).resolve() if args.output_path else project_path / "output.xlsx"
    controller = get_app_controller(str(project_path), args.log_level) # Исправленный вызов
    try:
        controller.export_project(str(output_path))
        print(f"Проект успешно экспортирован в {output_path}")
    except Exception as e:
        logger.error(f"Ошибка при экспорте проекта: {e}")
        sys.exit(1)

def process_data(args):
    """Обрабатывает данные в проекте."""
    logger.info(f"Обработка данных в проекте: {args.project_path}")
    project_path = Path(args.project_path).resolve()
    controller = get_app_controller(str(project_path), args.log_level) # Исправленный вызов
    try:
        controller.process_data()
        print(f"Данные в проекте {project_path} успешно обработаны")
    except Exception as e:
        logger.error(f"Ошибка при обработке данных: {e}")
        sys.exit(1)

def interactive_mode(args):
    """Запускает интерактивный режим (REPL)."""
    logger.info("Запуск интерактивного режима...")
    # Запускаем Python REPL с предзагруженными модулями
    # Это позволяет пользователю экспериментировать с API напрямую
    code_to_run = f"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path('{Path(__file__).resolve().parent}').resolve() / 'src'))

from utils.logger import setup_logger
from core.project_manager import ProjectManager
from core.app_controller import get_app_controller
from exceptions.app_exceptions import AppException

logger = setup_logger('__main__')

print("Добро пожаловать в интерактивный режим Excel Micro DB!")
print("Доступны: ProjectManager, get_app_controller, AppController, AppException, logger")
print("Пример: controller = get_app_controller('path/to/your/project')")
print("Введите 'exit()' для выхода.")
"""
    subprocess.run([sys.executable, "-c", code_to_run])

def main():
    parser = argparse.ArgumentParser(description="Excel Micro DB CLI")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="Уровень логирования")

    subparsers = parser.add_subparsers(dest="command", help="Доступные команды")

    # Подкоманда инициализации проекта
    parser_init = subparsers.add_parser("init", help="Инициализировать новый проект")
    parser_init.add_argument("--project-path", required=True, help="Путь для создания проекта")
    parser_init.set_defaults(func=init_project)

    # Подкоманда анализа файла
    parser_analyze = subparsers.add_parser("analyze", help="Анализировать Excel-файл")
    parser_analyze.add_argument("--file-path", required=True, help="Путь к анализируемому Excel-файлу")
    parser_analyze.add_argument("--project-path", required=True, help="Путь к проекту для сохранения результатов")
    parser_analyze.set_defaults(func=analyze_file)

    # Подкоманда экспорта проекта
    parser_export = subparsers.add_parser("export", help="Экспортировать проект в Excel")
    parser_export.add_argument("--project-path", required=True, help="Путь к проекту для экспорта")
    parser_export.add_argument("--output-path", help="Путь к результирующему Excel-файлу (по умолчанию: project_path/output.xlsx)")
    parser_export.set_defaults(func=export_project)

    # Подкоманда обработки данных
    parser_process = subparsers.add_parser("process", help="Обработать данные в проекте")
    parser_process.add_argument("--project-path", required=True, help="Путь к проекту для обработки")
    parser_process.set_defaults(func=process_data)

    # Подкоманда интерактивного режима
    parser_interactive = subparsers.add_parser("interactive", help="Запустить интерактивный режим (REPL)")
    parser_interactive.set_defaults(func=interactive_mode)

    # Если команда не указана, показываем справку
    if len(sys.argv) == 1:
        parser.print_help(sys.stderr)
        sys.exit(1)

    args = parser.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()