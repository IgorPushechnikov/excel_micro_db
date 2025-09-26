#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CLI точка входа в Excel Micro DB.
Поддерживает различные режимы работы через аргументы командной строки.
"""

import argparse
import sys
import os
import logging
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
    # --- НОВЫЙ КОД: Настройка лога в подпапке ---
    output_path_obj = Path(output_path)
    export_logs_dir = output_path_obj.parent / "logs"  # Создаём путь к подпапке logs в директории output
    export_logs_dir.mkdir(parents=True, exist_ok=True) # Создаём подпапку logs, если её нет

    log_file_path = export_logs_dir / f"export_{output_path_obj.stem}.log" # Имя файла лога

    # Создаём FileHandler
    export_file_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
    export_file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    export_file_handler.setFormatter(export_file_formatter)

    # Добавляем FileHandler к логгеру
    logger.addHandler(export_file_handler)
    logger.info(f"Экспорт результатов типа '{export_type}' в: {output_path}")
    # --- КОНЕЦ НОВОГО КОДА ---

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
            # --- НОВЫЙ КОД: Вызов дампа БД ---
            # Определяем путь к БД проекта
            project_db_path = Path(project_path) / "project_data.db"
            # Определяем папку для SQL-дампа (например, sql_export в той же папке, что и output)
            sql_export_dir = output_path_obj.parent / "sql_export"
            sql_export_dir.mkdir(parents=True, exist_ok=True) # Создаём подпапку sql_export, если её нет
            # Определяем путь к SQL-файлу
            sql_output_path = sql_export_dir / f"{project_db_path.name}.sql"
            logger.info(f"Начинается создание SQL-дампа БД в: {sql_output_path}")

            # Импортируем функцию dump_db_to_sql
            from src.utils.db_utils import dump_db_to_sql
            dump_success = dump_db_to_sql(str(project_db_path), str(sql_output_path))
            if dump_success:
                logger.info(f"SQL-дамп БД успешно создан: {sql_output_path}")
            else:
                logger.error(f"Не удалось создать SQL-дамп БД: {sql_output_path}")
            # --- КОНЕЦ НОВОГО КОДА ---
        else:
            logger.error("Экспорт через CLI завершился с ошибкой.")
            raise Exception("Ошибка экспорта")

    except Exception as e:
        logger.error(f"Ошибка при экспорте через CLI: {e}", exc_info=True) # exc_info=True для трассировки
        raise
    finally:
        # --- НОВЫЙ КОД: Удаляем FileHandler после завершения ---
        logger.removeHandler(export_file_handler)
        export_file_handler.close()
        # --- КОНЕЦ НОВОГО КОДА ---


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

def main() -> int:
    """Главная функция точки входа. Возвращает код завершения."""
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
  python main.py --gui
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
    
    mode_group.add_argument(
        '--gui',
        action='store_true',
        help='Запуск графического интерфейса пользователя (GUI)'
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
        if args.gui:
            # --- НОВОЕ: Запуск GUI ---
            logger.info("Запуск графического интерфейса...")
            # Импортируем и запускаем GUI
            try:
                from src.constructor.gui_app import main as gui_main
                logger.debug("Модуль GUI успешно импортирован.")
                # Передаём управление в GUI
                # gui_main() не принимает аргументы, как и CLI main()
                # Если нужно передать project_path из CLI, это нужно предусмотреть в gui_main
                # Пока просто запускаем
                return gui_main() # Возвращаем код завершения из GUI
            except ImportError as ie:
                logger.critical(f"Не удалось импортировать GUI: {ie}")
                print("Ошибка: Не удалось загрузить графический интерфейс. Убедитесь, что PySide6 установлен.")
                return 1
            except Exception as e_gui:
                logger.critical(f"Критическая ошибка при запуске GUI: {e_gui}", exc_info=True)
                print(f"Ошибка: Критическая ошибка при запуске GUI: {e_gui}")
                return 1
            # --- КОНЕЦ НОВОГО ---
            
        elif args.init:
            if not args.project_path:
                parser.error("--init требует указания --project-path")
            initialize_project(args.project_path)
            return 0 # Успех
            
        elif args.analyze:
            # --analyze теперь требует --project-path
            if not args.project_path:
                parser.error("--analyze требует указания --project-path")
            analyze_excel(args.analyze, args.project_path)
            return 0 # Успех
            
        elif args.export:
            # --export требует --output и --project-path
            if not args.output or not args.project_path:
                parser.error("--export требует указания --output и --project-path")
            export_results_cli(args.export, args.output, args.project_path)
            return 0 # Успех
            
        elif args.process:
            if not args.config:
                parser.error("--process требует указания --config")
            process_data(args.config)
            return 0 # Успех
            
        elif args.interactive:
            start_interactive_mode()
            return 0 # Успех
            
    except Exception as e:
        logger.critical(f"Критическая ошибка приложения: {e}")
        return 1

    # Если мы здесь, значит, ни один из режимов не был выбран или argparse обработал help/version.
    # Это маловероятно из-за required=True, но для строгой типизации добавим явный return.
    return 0 # Считаем успешным завершением (например, вывод справки)

if __name__ == "__main__":
    # main() возвращает код, который мы передаём в sys.exit()
    exit_code = main() 
    # Если main() не вернул ничего (None), sys.exit() интерпретирует это как 0
    sys.exit(exit_code)