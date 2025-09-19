#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CLI точка входа в Excel Micro DB.
Поддерживает различные режимы работы через аргументы командной строки.
"""

import argparse
import sys
import os
from pathlib import Path

# Добавляем src в путь поиска модулей
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.utils.logger import get_logger
# Импортируем AppController для интеграции
from src.core.app_controller import create_app_controller
# ProjectManager больше не импортируем напрямую, так как AppController его использует

# Получаем логгер для этого модуля
logger = get_logger(__name__)

def initialize_project(project_path: str) -> None:
    """Инициализация нового проекта."""
    logger.info(f"Инициализация проекта в директории: {project_path}")
    try:
        # Используем AppController для создания проекта
        # Это обеспечивает единую точку входа и правильную инициализацию всех компонентов
        app_controller = create_app_controller()
        if not app_controller.initialize():
             logger.error("Не удалось инициализировать приложение для создания проекта.")
             raise Exception("Ошибка инициализации приложения")
             
        success = app_controller.create_project(project_path)
        
        if success:
            logger.info("Проект инициализирован успешно")
        else:
            logger.error("Не удалось инициализировать проект")
            raise Exception("Ошибка создания проекта")
            
    except Exception as e:
        logger.error(f"Ошибка при инициализации проекта: {e}")
        raise

def analyze_excel(file_path: str, project_path: str) -> None:
    """Анализ Excel файла через AppController."""
    logger.info(f"Начало анализа файла: {file_path}")
    try:
        # Проверка существования файла
        if not Path(file_path).exists():
            logger.error(f"Файл не найден: {file_path}")
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        # Создаем и инициализируем контроллер приложения
        app_controller = create_app_controller(project_path=project_path)
        if not app_controller.initialize():
             logger.error("Не удалось инициализировать приложение.")
             raise Exception("Ошибка инициализации приложения")

        # Проверяем, загружен ли проект (initialize должен был это сделать)
        if not app_controller.is_project_loaded:
             logger.error("Проект не загружен. Убедитесь, что путь к проекту корректен.")
             raise Exception("Проект не загружен")

        # Вызываем анализ через контроллер
        # Передаем опции анализа, если нужно, например:
        options = {
            'max_rows': 1000,
            'include_formulas': True
        }
        success = app_controller.analyze_excel_file(file_path, options=options)

        if success:
            logger.info("Анализ файла завершен успешно")
        else:
            logger.error("Анализ файла завершился с ошибкой")
            raise Exception("Ошибка анализа файла")

    except Exception as e:
        logger.error(f"Ошибка при анализе файла: {e}")
        raise # Повторно вызываем исключение, чтобы main() мог его обработать

def process_data(config_path: str) -> None:
    """Обработка данных по конфигурации."""
    logger.info(f"Обработка данных с конфигурацией: {config_path}")
    try:
        # TODO: Здесь будет вызов процессора через AppController
        # Пока только заглушка
        if not Path(config_path).exists():
            logger.error(f"Конфигурационный файл не найден: {config_path}")
            return
            
        logger.info("Обработка данных завершена")
    except Exception as e:
        logger.error(f"Ошибка при обработке данных: {e}")
        raise

def export_results_cli(export_type: str, output_path: str, project_path: str) -> None:
    """Экспорт результатов проекта через CLI."""
    logger.info(f"Экспорт результатов типа '{export_type}' в: {output_path}")
    try:
        logger.debug(f"[export_results_cli] Создание AppController с project_path: {project_path}")
        # Создаем и инициализируем контроллер
        app_controller = create_app_controller(project_path=project_path)
        logger.debug("[export_results_cli] Вызов app_controller.initialize()")
        init_success = app_controller.initialize()
        if not init_success:
            logger.error("Не удалось инициализировать приложение для экспорта.")
            raise Exception("Ошибка инициализации приложения")
        logger.debug("[export_results_cli] Инициализация прошла успешно")

        # Проверим, загружен ли проект.
        logger.debug(f"[export_results_cli] Статус проекта: is_project_loaded={app_controller.is_project_loaded}")
        if not app_controller.is_project_loaded:
             logger.error("Проект не загружен. Убедитесь, что путь к проекту корректен.")
             raise Exception("Проект не загружен")
        logger.debug("[export_results_cli] Проект загружен, переходим к экспорту")

        # Вызываем экспорт через контроллер
        logger.debug(f"[export_results_cli] Вызов app_controller.export_results(type={export_type}, path={output_path})")
        success = app_controller.export_results(export_type=export_type, output_path=output_path)
        logger.debug(f"[export_results_cli] app_controller.export_results вернул: {success}")

        if success:
            logger.info("Экспорт через CLI завершен успешно.")
        else:
            logger.error("Экспорт через CLI завершился с ошибкой.")
            raise Exception("Ошибка экспорта")

    except Exception as e:
        logger.error(f"Ошибка при экспорте через CLI: {e}", exc_info=True) # exc_info=True для трассировки
        raise


def start_interactive_mode() -> None:
    """Запуск интерактивного режима (REPL)."""
    logger.info("Запуск интерактивного режима")
    print("Добро пожаловать в интерактивный режим Excel Micro DB!")
    print("Для выхода введите 'exit' или нажмите Ctrl+C")
    
    # TODO: Здесь будет реализация интерактивного режима
    # Пока только демонстрация
    try:
        while True:
            try:
                command = input("excel_micro_db> ").strip()
                if command.lower() in ['exit', 'quit']:
                    break
                elif command:
                    print(f"Выполняется команда: {command}")
                    # Здесь будет обработка команд
            except KeyboardInterrupt:
                print("\nПолучен сигнал завершения (Ctrl+C)")
                break
        
        logger.info("Выход из интерактивного режима")
    except Exception as e:
        logger.error(f"Ошибка в интерактивном режиме: {e}")
        raise

def main():
    """Главная функция точки входа."""
    parser = argparse.ArgumentParser(
        description="Excel Micro DB - микро-СУБД для анализа и воссоздания логики Excel файлов",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python main.py --init --project-path ./my_project
  python main.py --analyze ./data/input.xlsx --project-path ./my_project
  python main.py --export excel --output ./output/result.xlsx --project-path ./my_project
  python main.py --process --config config/batch.yaml
  python main.py --interactive
        """
    )
    
    # Группы взаимоисключающих аргументов (режимы работы)
    # Обновляем группу, чтобы включить --export
    mode_group = parser.add_mutually_exclusive_group(required=True)
    
    mode_group.add_argument(
        '--init',
        action='store_true',
        help='Инициализация нового проекта'
    )
    
    mode_group.add_argument(
        '--analyze',
        metavar='FILE',
        help='Анализ Excel файла (требует --project-path)'
    )
    
    # Добавляем новый режим экспорта
    mode_group.add_argument(
        '--export',
        metavar='TYPE',
        choices=['excel'], # Ограничиваем поддерживаемые типы на данном этапе
        help='Экспорт результатов проекта (например, excel). Требует --output и --project-path.'
    )
    
    mode_group.add_argument(
        '--process',
        action='store_true',
        help='Обработка данных'
    )
    
    mode_group.add_argument(
        '--interactive',
        action='store_true',
        help='Запуск интерактивного режима (REPL)'
    )
    
    # Дополнительные аргументы
    parser.add_argument(
        '--project-path',
        metavar='PATH',
        help='Путь к директории проекта'
    )
    
    parser.add_argument(
        '--config',
        metavar='FILE',
        help='Путь к конфигурационному файлу (для --process)'
    )
    
    # Добавляем аргумент для выходного файла экспорта
    parser.add_argument(
        '--output',
        metavar='FILE',
        help='Путь к выходному файлу (для --export)'
    )
    
    # Парсим аргументы
    args = parser.parse_args()
    
    try:
        # Обработка выбранных режимов
        if args.init:
            if not args.project_path:
                parser.error("--init требует указания --project-path")
            initialize_project(args.project_path)
            
        elif args.analyze:
            # --analyze теперь требует --project-path
            if not args.project_path:
                parser.error("--analyze требует указания --project-path")
            analyze_excel(args.analyze, args.project_path)
            
        elif args.export:
            # --export требует --output и --project-path
            if not args.output or not args.project_path:
                parser.error("--export требует указания --output и --project-path")
            export_results_cli(args.export, args.output, args.project_path)
            
        elif args.process:
            if not args.config:
                parser.error("--process требует указания --config")
            process_data(args.config)
            
        elif args.interactive:
            start_interactive_mode()
            
    except Exception as e:
        logger.critical(f"Критическая ошибка приложения: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()