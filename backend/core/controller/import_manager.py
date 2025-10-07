# backend/core/controller/import_manager.py

import logging
from typing import Dict, Any, List, Optional, Callable
from backend.storage.base import ProjectDBStorage  # <-- Импортируем ProjectDBStorage
from backend.utils.logger import get_logger
# Импортируем функции импорта из app_controller_data_import
from ..app_controller_data_import import (
    import_raw_data_from_excel,
    import_styles_from_excel,
    import_charts_from_excel,
    import_formulas_from_excel,
    import_raw_data_from_excel_in_chunks,
    # import_raw_data_from_excel_selective, # Добавить при необходимости
    # import_styles_from_excel_selective, # Добавить при необходимости
    # и т.д.
    # --- НОВОЕ: Импорт функции для "сырых" значений ---
    import_raw_values_only_from_excel # <-- НОВОЕ
)

logger = get_logger(__name__)

class ImportManager:
    def __init__(self, app_controller):
        """
        Инициализирует ImportManager.
        Не хранит ссылку на storage, создаёт его в потоке метода.

        Args:
            app_controller: Экземпляр AppController.
        """
        self.app_controller = app_controller
        logger.debug("ImportManager инициализирован.")

    def perform_import_raw_data(self, file_path: str, db_path: str, progress_callback: Optional[Callable[[int, str], None]] = None, options: Optional[Dict[str, Any]] = None) -> bool: # <-- ИЗМЕНЕНО: Добавлен options
        """
        Выполняет импорт "сырых" данных (значения, формулы как строки).
        Создаёт собственное соединение с БД в текущем потоке.

        Args:
            file_path (str): Путь к Excel-файлу.
            db_path (str): Путь к файлу БД проекта (.db).
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
            options (Optional[Dict[str, Any]]): Опции импорта.
                {
                    'sheets': List[str],     # Список имён листов для импорта.
                    'chunk_size_rows': int, # Количество строк в одной части (по умолчанию 100)
                    # Другие опции в будущем...
                }

        Returns:
            bool: True, если импорт успешен.
        """
        storage = ProjectDBStorage(db_path)
        if not storage.connect():
            logger.error(f"ImportManager: Не удалось подключиться к БД проекта {db_path}.")
            return False

        try:
            logger.info(f"ImportManager: Начало импорта 'сырых' данных из {file_path}.")

            if progress_callback:
                progress_callback(0, f"Импорт 'сырых' данных из {file_path}...")

            # --- ИЗМЕНЕНО: Передаём options ---
            success = import_raw_data_from_excel(storage, file_path, options=options)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            if progress_callback:
                progress_callback(100 if success else 0, f"Импорт 'сырых' данных {'завершён' if success else 'не удался'}.")

            if success:
                logger.info(f"ImportManager: Импорт 'сырых' данных из {file_path} завершён успешно.")
            else:
                logger.error(f"ImportManager: Ошибка импорта 'сырых' данных из {file_path}.")
            return success

        except Exception as e:
            logger.error(f"ImportManager: Ошибка при импорте 'сырых' данных из {file_path}: {e}", exc_info=True)
            if progress_callback:
                progress_callback(0, f"Ошибка импорта 'сырых' данных: {e}")
            return False
        finally:
            storage.disconnect()
            logger.debug(f"ImportManager: Соединение с БД {db_path} закрыто.")

    def perform_import_raw_values_only(self, file_path: str, db_path: str, progress_callback: Optional[Callable[[int, str], None]] = None, options: Optional[Dict[str, Any]] = None) -> bool: # <-- ИЗМЕНЕНО: Добавлен options
        """
        Выполняет импорт "сырых" значений (только результаты формул и значения ячеек).
        Создаёт собственное соединение с БД в текущем потоке.

        Args:
            file_path (str): Путь к Excel-файлу.
            db_path (str): Путь к файлу БД проекта (.db).
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
            options (Optional[Dict[str, Any]]): Опции импорта.
                {
                    'sheets': List[str],     # Список имён листов для импорта.
                    'chunk_size_rows': int, # Количество строк в одной части (по умолчанию 100)
                    # Другие опции в будущем...
                }

        Returns:
            bool: True, если импорт успешен.
        """
        # --- ИЗМЕНЕНО: Импортируем функцию напрямую ---
        # from ..app_controller_data_import import import_raw_values_only_from_excel as import_func # <-- УДАЛЕНО
        # --- КОНЕЦ ИЗМЕНЕНИЯ ---

        storage = ProjectDBStorage(db_path)
        if not storage.connect():
            logger.error(f"ImportManager: Не удалось подключиться к БД проекта {db_path}.")
            return False

        try:
            logger.info(f"ImportManager: Начало импорта 'сырых' значений (только результаты) из {file_path}.")

            if progress_callback:
                progress_callback(0, f"Импорт 'сырых' значений (только результаты) из {file_path}...")

            # --- ИЗМЕНЕНО: Передаём options ---
            success = import_raw_values_only_from_excel(storage, file_path, options=options)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            if progress_callback:
                progress_callback(100 if success else 0, f"Импорт 'сырых' значений (только результаты) {'завершён' if success else 'не удался'}.")

            if success:
                logger.info(f"ImportManager: Импорт 'сырых' значений (только результаты) из {file_path} завершён успешно.")
            else:
                logger.error(f"ImportManager: Ошибка импорта 'сырых' значений (только результаты) из {file_path}.")
            return success

        # Обработка специфичной ошибки openpyxl
        except TypeError as e:
            if "Nested.from_tree()" in str(e) and "missing 1 required positional argument" in str(e):
                logger.error(f"ImportManager: Ошибка openpyxl при импорте 'сырых' значений (только результаты) из {file_path}. Файл может содержать сложные структуры. Подробности: {e}", exc_info=True)
                if progress_callback:
                    progress_callback(0, f"Ошибка openpyxl: {e}. Файл содержит несовместимые структуры.")
                return False
            else:
                raise
        except Exception as e:
            logger.error(f"ImportManager: Ошибка при импорте 'сырых' значений (только результаты) из {file_path}: {e}", exc_info=True)
            if progress_callback:
                progress_callback(0, f"Ошибка импорта 'сырых' значений (только результаты): {e}")
            return False
        finally:
            storage.disconnect()
            logger.debug(f"ImportManager: Соединение с БД {db_path} закрыто.")

    def perform_import_styles(self, file_path: str, db_path: str, progress_callback: Optional[Callable[[int, str], None]] = None, options: Optional[Dict[str, Any]] = None) -> bool: # <-- ИЗМЕНЕНО: Добавлен options
        """
        Выполняет импорт стилей.
        Создаёт собственное соединение с БД в текущем потоке.

        Args:
            file_path (str): Путь к Excel-файлу.
            db_path (str): Путь к файлу БД проекта (.db).
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
            options (Optional[Dict[str, Any]]): Опции импорта.
                {
                    'sheets': List[str],     # Список имён листов для импорта.
                    'chunk_size_rows': int, # Количество строк в одной части (по умолчанию 100)
                    # Другие опции в будущем...
                }

        Returns:
            bool: True, если импорт успешен.
        """
        storage = ProjectDBStorage(db_path)
        if not storage.connect():
            logger.error(f"ImportManager: Не удалось подключиться к БД проекта {db_path}.")
            return False

        try:
            logger.info(f"ImportManager: Начало импорта стилей из {file_path}.")

            if progress_callback:
                progress_callback(0, f"Импорт стилей из {file_path}...")

            # --- ИЗМЕНЕНО: Передаём options ---
            success = import_styles_from_excel(storage, file_path, options=options)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            if progress_callback:
                progress_callback(100 if success else 0, f"Импорт стилей {'завершён' if success else 'не удался'}.")

            if success:
                logger.info(f"ImportManager: Импорт стилей из {file_path} завершён успешно.")
            else:
                logger.error(f"ImportManager: Ошибка импорта стилей из {file_path}.")
            return success

        except Exception as e:
            logger.error(f"ImportManager: Ошибка при импорте стилей из {file_path}: {e}", exc_info=True)
            if progress_callback:
                progress_callback(0, f"Ошибка импорта стилей: {e}")
            return False
        finally:
            storage.disconnect()
            logger.debug(f"ImportManager: Соединение с БД {db_path} закрыто.")

    def perform_import_charts(self, file_path: str, db_path: str, progress_callback: Optional[Callable[[int, str], None]] = None, options: Optional[Dict[str, Any]] = None) -> bool: # <-- ИЗМЕНЕНО: Добавлен options
        """
        Выполняет импорт диаграмм.
        Создаёт собственное соединение с БД в текущем потоке.

        Args:
            file_path (str): Путь к Excel-файлу.
            db_path (str): Путь к файлу БД проекта (.db).
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
            options (Optional[Dict[str, Any]]): Опции импорта.
                {
                    'sheets': List[str],     # Список имён листов для импорта.
                    # 'chunk_size_charts': int, # Количество диаграмм в одной части (по умолчанию len(all_charts_on_sheet))
                    # Другие опции в будущем...
                }

        Returns:
            bool: True, если импорт успешен.
        """
        storage = ProjectDBStorage(db_path)
        if not storage.connect():
            logger.error(f"ImportManager: Не удалось подключиться к БД проекта {db_path}.")
            return False

        try:
            logger.info(f"ImportManager: Начало импорта диаграмм из {file_path}.")

            if progress_callback:
                progress_callback(0, f"Импорт диаграмм из {file_path}...")

            # --- ИЗМЕНЕНО: Передаём options ---
            success = import_charts_from_excel(storage, file_path, options=options)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            if progress_callback:
                progress_callback(100 if success else 0, f"Импорт диаграмм {'завершён' if success else 'не удался'}.")

            if success:
                logger.info(f"ImportManager: Импорт диаграмм из {file_path} завершён успешно.")
            else:
                logger.error(f"ImportManager: Ошибка импорта диаграмм из {file_path}.")
            return success

        except Exception as e:
            logger.error(f"ImportManager: Ошибка при импорте диаграмм из {file_path}: {e}", exc_info=True)
            if progress_callback:
                progress_callback(0, f"Ошибка импорта диаграмм: {e}")
            return False
        finally:
            storage.disconnect()
            logger.debug(f"ImportManager: Соединение с БД {db_path} закрыто.")

    def perform_import_formulas(self, file_path: str, db_path: str, progress_callback: Optional[Callable[[int, str], None]] = None, options: Optional[Dict[str, Any]] = None) -> bool: # <-- ИЗМЕНЕНО: Добавлен options
        """
        Выполняет импорт формул.
        Создаёт собственное соединение с БД в текущем потоке.

        Args:
            file_path (str): Путь к Excel-файлу.
            db_path (str): Путь к файлу БД проекта (.db).
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
            options (Optional[Dict[str, Any]]): Опции импорта.
                {
                    'sheets': List[str],     # Список имён листов для импорта.
                    'chunk_size_rows': int, # Количество строк в одной части (по умолчанию 100)
                    # Другие опции в будущем...
                }

        Returns:
            bool: True, если импорт успешен.
        """
        storage = ProjectDBStorage(db_path)
        if not storage.connect():
            logger.error(f"ImportManager: Не удалось подключиться к БД проекта {db_path}.")
            return False

        try:
            logger.info(f"ImportManager: Начало импорта формул из {file_path}.")

            if progress_callback:
                progress_callback(0, f"Импорт формул из {file_path}...")

            # --- ИЗМЕНЕНО: Передаём options ---
            success = import_formulas_from_excel(storage, file_path, options=options)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            if progress_callback:
                progress_callback(100 if success else 0, f"Импорт формул {'завершён' if success else 'не удался'}.")

            if success:
                logger.info(f"ImportManager: Импорт формул из {file_path} завершён успешно.")
            else:
                logger.error(f"ImportManager: Ошибка импорта формул из {file_path}.")
            return success

        except Exception as e:
            logger.error(f"ImportManager: Ошибка при импорте формул из {file_path}: {e}", exc_info=True)
            if progress_callback:
                progress_callback(0, f"Ошибка импорта формул: {e}")
            return False
        finally:
            storage.disconnect()
            logger.debug(f"ImportManager: Соединение с БД {db_path} закрыто.")

    def perform_import_raw_data_in_chunks(self, file_path: str, db_path: str, progress_callback: Optional[Callable[[int, str], None]] = None, options: Optional[Dict[str, Any]] = None) -> bool: # <-- ИЗМЕНЕНО: Добавлен options
        """
        Выполняет импорт "сырых" данных частями.
        Создаёт собственное соединение с БД в текущем потоке.

        Args:
            file_path (str): Путь к Excel-файлу.
            db_path (str): Путь к файлу БД проекта (.db).
            progress_callback (Optional[Callable[[int, str], None]]): Функция для обновления прогресса.
            options (Optional[Dict[str, Any]]): Опции импорта.
                {
                    'sheets': List[str],     # Список имён листов для импорта.
                    'chunk_size_rows': int, # Количество строк в одной части (по умолчанию 100)
                    # Другие опции в будущем...
                }

        Returns:
            bool: True, если импорт успешен.
        """
        storage = ProjectDBStorage(db_path)
        if not storage.connect():
            logger.error(f"ImportManager: Не удалось подключиться к БД проекта {db_path}.")
            return False

        try:
            logger.info(f"ImportManager: Начало импорта 'сырых' данных частями из {file_path}.")

            if progress_callback:
                progress_callback(0, f"Импорт 'сырых' данных частями из {file_path}...")

            # --- ИЗМЕНЕНО: Передаём options как chunk_options ---
            # chunk_options можно передать сюда из AppController, если нужно
            # Используем пустой словарь, если chunk_options не предоставлены
            # success = import_raw_data_from_excel_in_chunks(storage, file_path, chunk_options={}) # <-- СТАРОЕ
            success = import_raw_data_from_excel_in_chunks(storage, file_path, chunk_options=options or {}) # <-- НОВОЕ
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            if progress_callback:
                progress_callback(100 if success else 0, f"Импорт 'сырых' данных частями {'завершён' if success else 'не удался'}.")

            if success:
                logger.info(f"ImportManager: Импорт 'сырых' данных частями из {file_path} завершён успешно.")
            else:
                logger.error(f"ImportManager: Ошибка импорта 'сырых' данных частями из {file_path}.")
            return success

        except Exception as e:
            logger.error(f"ImportManager: Ошибка при импорте 'сырых' данных частями из {file_path}: {e}", exc_info=True)
            if progress_callback:
                progress_callback(0, f"Ошибка импорта 'сырых' данных частями: {e}")
            return False
        finally:
            storage.disconnect()
            logger.debug(f"ImportManager: Соединение с БД {db_path} закрыто.")
