# src/core/app_controller.py

"""
Основной контроллер приложения Excel Micro DB.

Координирует работу всех компонентов системы.
"""

import sys
from pathlib import Path
from typing import Optional, Dict, Any, TYPE_CHECKING

# Добавляем корень проекта в путь поиска модулей если нужно
project_root = Path(__file__).parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.utils.logger import get_logger
from src.core.project_manager import ProjectManager

# --- ИНТЕГРАЦИЯ STORAGE: Импорт модуля хранения ---
# ИСПРАВЛЕНО: Импорт из правильного модуля
# from src.storage.database import ProjectDBStorage # <-- СТАРОЕ
from src.storage.base import ProjectDBStorage    # <-- НОВОЕ

# Импорт для аннотаций типов, избегая циклических импортов
if TYPE_CHECKING:
    from src.constructor.widgets.sheet_editor import SheetEditor  # type: ignore

# --- ИНТЕГРАЦИЯ ANALYZER: Импорт модуля анализа ---
from src.analyzer.logic_documentation import analyze_excel_file

# Получаем логгер для этого модуля
logger = get_logger(__name__)


class AppController:
    """
    Основной контроллер приложения Excel Micro DB.

    Отвечает за:
    - Инициализацию и управление компонентами приложения
    - Координацию между модулями (analyzer, processor, storage, exporter, constructor)
    - Управление состоянием приложения
    - Обработку ошибок на уровне приложения
    """

    def __init__(self, project_path: Optional[str] = None):
        """
        Инициализация контроллера приложения.

        Args:
            project_path (Optional[str]): Путь к директории проекта
        """
        logger.debug("Инициализация AppController")
        self.project_path = Path(project_path) if project_path else None
        self.project_manager: Optional[ProjectManager] = None
        self.is_initialized = False

        # Состояние приложения
        self.current_project: Optional[Dict[str, Any]] = None
        self.is_project_loaded = False

        logger.debug("AppController инициализирован")

    def initialize(self) -> bool:
        """
        Инициализация приложения и его компонентов.

        Returns:
            bool: True если инициализация успешна, False в противном случае
        """
        try:
            logger.info("Начало инициализации приложения")

            # Инициализация менеджера проектов
            self.project_manager = ProjectManager()

            # Если указан путь к проекту, загружаем его (без рекурсивного вызова initialize)
            if self.project_path and self.project_path.exists():
                # Прямая загрузка без проверки is_initialized
                logger.info(f"Загрузка проекта из: {self.project_path}")
                project_data = self.project_manager.load_project(str(self.project_path))
                if project_data:
                    self.current_project = project_data
                    self.is_project_loaded = True
                    logger.info("Проект загружен успешно")
                else:
                    logger.error("Не удалось загрузить проект")

            self.is_initialized = True
            logger.info("Инициализация приложения завершена успешно")
            return True

        except Exception as e:
            logger.error(f"Ошибка при инициализации приложения: {e}")
            self.is_initialized = False
            return False

    def create_project(self, project_path: str, project_name: Optional[str] = None) -> bool:
        """
        Создание нового проекта.

        Args:
            project_path (str): Путь к директории проекта
            project_name (Optional[str]): Название проекта (по умолчанию имя директории)

        Returns:
            bool: True если проект создан успешно, False в противном случае
        """
        if not self.is_initialized:
            logger.warning("Приложение не инициализировано. Выполняем инициализацию...")
            if not self.initialize():
                return False

        try:
            logger.info(f"Создание нового проекта в: {project_path}")

            if not self.project_manager:
                logger.error("Менеджер проектов не инициализирован")
                return False

            success = self.project_manager.create_project(project_path, project_name)

            if success:
                logger.info("Проект создан успешно")
                # Автоматически загружаем созданный проект
                self.load_project(project_path)
                return True
            else:
                logger.error("Не удалось создать проект")
                return False

        except Exception as e:
            logger.error(f"Ошибка при создании проекта: {e}")
            return False

    def load_project(self, project_path: str) -> bool:
        """
        Загрузка существующего проекта.

        Args:
            project_path (str): Путь к директории проекта

        Returns:
            bool: True если проект загружен успешно, False в противном случае
        """
        if not self.is_initialized:
            logger.warning("Приложение не инициализировано. Выполняем инициализацию...")
            if not self.initialize():
                return False

        try:
            logger.info(f"Загрузка проекта из: {project_path}")

            if not self.project_manager:
                logger.error("Менеджер проектов не инициализирован")
                return False

            project_data = self.project_manager.load_project(project_path)

            if project_data:
                self.current_project = project_data
                self.project_path = Path(project_path)
                self.is_project_loaded = True
                logger.info("Проект загружен успешно")
                return True
            else:
                logger.error("Не удалось загрузить проект")
                return False

        except Exception as e:
            logger.error(f"Ошибка при загрузке проекта: {e}")
            return False

    def analyze_excel_file(self, file_path: str, options: Optional[Dict[str, Any]] = None) -> bool:
        """
        Анализ Excel файла.

        Args:
            file_path (str): Путь к Excel файлу для анализа
            options (Optional[Dict[str, Any]]): Дополнительные опции анализа

        Returns:
            bool: True если анализ выполнен успешно, False в противном случае
        """
        if not self.is_project_loaded:
            logger.warning("Проект не загружен. Невозможно выполнить анализ")
            return False

        if not self.project_path:
            logger.error("Путь к проекту не установлен.")
            return False

        try:
            logger.info(f"Начало анализа файла: {file_path}")

            # Проверка существования файла
            if not Path(file_path).exists():
                logger.error(f"Файл не найден: {file_path}")
                return False

            # - ИНТЕГРАЦИЯ ANALYZER: Вызов анализатора -
            # Передаем опции анализа
            documentation_data = analyze_excel_file(file_path)
            if documentation_data is None:
                logger.error("Анализатор вернул None. Ошибка при анализе файла.")
                return False

            logger.info("Анализ файла завершен успешно")

            # --- НАЧАЛО ИЗМЕНЕНИЙ ---
            # Получаем имя проекта для сохранения результатов
            # ИСПРАВЛЕНО: Проверка на None
            if self.current_project:
                project_name = self.current_project.get("project_name", self.project_path.name if self.project_path else "UnknownProject")
            else:
                project_name = self.project_path.name if self.project_path else "UnknownProject"
            logger.info("Начало сохранения результатов анализа в хранилище")

            # Определяем путь к файлу БД
            db_path = self.project_path / "project_data.db"
            logger.debug(f"Путь к БД проекта: {db_path}")

            # Используем контекстный менеджер для автоматического управления соединением
            # ИСПРАВЛЕНО: Импорт
            # from src.storage.database import ProjectDBStorage # <-- СТАРОЕ
            from src.storage.base import ProjectDBStorage    # <-- НОВОЕ

            try:
                with ProjectDBStorage(str(db_path)) as storage:  # <-- connect() вызывается автоматически
                    logger.debug("Соединение с БД установлено (через контекстный менеджер)")
                    # Сохраняем результаты анализа
                    save_success = storage.save_analysis_results(project_name, documentation_data)
                # После выхода из блока 'with' соединение автоматически закрывается (disconnect())

                if save_success:
                    logger.info("Результаты анализа успешно сохранены в хранилище")
                    return True
                else:
                    logger.error("Ошибка при сохранении результатов анализа в хранилище")
                    return False
            except Exception as e:
                logger.error(f"Ошибка при работе с БД проекта: {e}", exc_info=True)
                return False
        except Exception as e:
            logger.error(f"Ошибка при анализе файла или сохранении в хранилище: {e}", exc_info=True)
            return False

    def process_data(self, config_path: str) -> bool:
        """
        Обработка данных по конфигурации.

        Args:
            config_path (str): Путь к конфигурационному файлу обработки

        Returns:
            bool: True если обработка выполнена успешно, False в противном случае
        """
        if not self.is_project_loaded:
            logger.warning("Проект не загружен. Невозможно выполнить обработку")
            return False

        try:
            logger.info(f"Начало обработки данных с конфигурацией: {config_path}")

            # TODO: Здесь будет интеграция с процессором

            # Пока только демонстрация логики
            if not Path(config_path).exists():
                logger.error(f"Конфигурационный файл не найден: {config_path}")
                return False

            # Здесь будет вызов модуля обработки
            # from src.processor.data_processor import process_data_with_config
            # result = process_data_with_config(config_path)

            logger.info("Обработка данных завершена успешно")
            return True

        except Exception as e:
            logger.error(f"Ошибка при обработке данных: {e}")
            return False

    def export_results(self, export_type: str, output_path: str) -> bool:
        """
        Экспорт результатов.

        Args:
            export_type (str): Тип экспорта (например, 'documentation', 'report', 'data', 'excel').
                               Пока поддерживается только 'excel'.
            output_path (str): Путь для сохранения результата.

        Returns:
            bool: True если экспорт выполнен успешно, False в противном случае.
        """
        if not self.is_project_loaded:
            logger.warning("Проект не загружен. Невозможно выполнить экспорт.")
            return False

        if not self.project_path:
            logger.error("Путь к проекту не установлен.")
            return False

        # Пока поддерживаем только экспорт в Excel
        if export_type.lower() != 'excel':
            logger.warning(f"Тип экспорта '{export_type}' пока не поддерживается. Поддерживается только 'excel'.")
            # TODO: Здесь можно добавить поддержку других типов экспорта позже
            return False

        try:
            logger.info(f"Начало экспорта проекта в Excel: {output_path}")

            # --- ИНТЕГРАЦИЯ EXPORTER ---
            # Импортируем функцию экспорта прямо здесь, чтобы избежать циклических импортов
            # на уровне модуля, если exporter когда-нибудь будет импортировать AppController.
            # ИСПРАВЛЕНО: Импорт из правильного модуля и правильной функции
            # from src.exporter.excel_exporter import export_project_to_excel # <-- СТАРОЕ
            from src.exporter.direct_db_exporter import export_project_from_db # <-- НОВОЕ

            # Вызываем функцию экспорта, передавая путь к проекту и путь к выходному файлу
            # Определяем путь к БД проекта
            db_path = self.project_path / "project_data.db"
            # ИСПРАВЛЕНО: Вызов правильной функции с правильными аргументами
            # success = export_project_to_excel(...) # <-- СТАРОЕ
            success = export_project_from_db( # <-- НОВОЕ
                db_path=str(db_path),
                output_path=output_path
                # project_path убран, так как не нужен
            )

            if success:
                logger.info("Экспорт в Excel завершен успешно.")
                return True
            else:
                logger.error("Ошибка при экспорте в Excel.")
                return False

        except Exception as e:
            logger.error(f"Ошибка при экспорте результатов: {e}", exc_info=True)
            return False

    # === НОВОЕ: Метод для получения редактируемых данных листа ===

    def get_sheet_editable_data(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """
        Получает редактируемые данные листа из БД проекта.

        Вызывается SheetEditor для загрузки содержимого.

        Args:
            sheet_name (str): Имя листа Excel.

        Returns:
            Optional[Dict[str, Any]]: Словарь с ключами 'column_names' и 'rows',
                                      или None в случае ошибки или если проект не загружен.
        """
        if not self.is_project_loaded:
            logger.warning("Проект не загружен. Невозможно получить данные листа.")
            return None

        if not self.project_path:
            logger.error("Путь к проекту не установлен.")
            return None

        try:
            db_path = self.project_path / "project_data.db"
            logger.debug(f"AppController: Получение редактируемых данных для листа '{sheet_name}' из БД: {db_path}")

            with ProjectDBStorage(str(db_path)) as storage:
                editable_data = storage.load_sheet_editable_data(sheet_name)

                if editable_data and 'column_names' in editable_data:
                    logger.info(f"Редактируемые данные для листа '{sheet_name}' успешно получены.")
                    return editable_data
                else:
                    logger.warning(f"Редактируемые данные для листа '{sheet_name}' не найдены или пусты.")
                    return {"column_names": [], "rows": []}  # Возвращаем пустую структуру

        except Exception as e:
            logger.error(f"Ошибка при получении редактируемых данных для листа '{sheet_name}': {e}", exc_info=True)
            return None

    # =========================================================

    # === НОВОЕ: Метод для обновления ячейки и сохранения истории ===

    def update_sheet_cell_in_project(self, sheet_name: str, row_index: int, column_name: str, new_value: str) -> bool:
        """
        Обновляет значение ячейки в редактируемых данных проекта и сохраняет запись в истории.

        Вызывается SheetEditor при редактировании ячейки.

        Args:
            sheet_name (str): Имя листа Excel.
            row_index (int): 0-базовый индекс строки.
            column_name (str): Имя столбца.
            new_value (str): Новое значение ячейки.

        Returns:
            bool: True если обновление и сохранение истории прошли успешно, False в противном случае.
        """
        if not self.is_project_loaded:
            logger.warning("Проект не загружен. Невозможно обновить ячейку.")
            return False

        if not self.project_path:
            logger.error("Путь к проекту не установлен.")
            return False

        try:
            db_path = self.project_path / "project_data.db"
            logger.debug(f"AppController: Обновление ячейки [{sheet_name}][{row_index}, {column_name}] в БД: {db_path}")

            with ProjectDBStorage(str(db_path)) as storage:
                conn = storage.connection
                if not conn:
                    logger.error("Нет активного соединения с БД в AppController.update_sheet_cell_in_project")
                    return False

                cursor = conn.cursor()
                cursor.execute("SELECT id FROM sheets WHERE name = ?", (sheet_name,))
                sheet_row = cursor.fetchone()
                if not sheet_row:
                    logger.error(f"Лист '{sheet_name}' не найден в БД при попытке обновления ячейки.")
                    return False
                sheet_id = sheet_row[0]
                logger.debug(f"AppController: Найден sheet_id={sheet_id} для листа '{sheet_name}'.")

                # --- ИЗМЕНЕНИЯ НАЧИНАЮТСЯ ЗДЕСЬ ---
                
                # Импортируем вспомогательные функции
                from src.storage.base import sanitize_table_name, sanitize_column_name

                # 2. Получить старое значение для истории
                # Мы можем получить его напрямую из таблицы редактируемых данных
                # ИСПРАВЛЕНО: Вызов функции как функции, а не метода объекта
                editable_table_name = f"editable_data_{sanitize_table_name(sheet_name)}"

                # Санитизируем имя столбца
                # ИСПРАВЛЕНО: Вызов функции как функции, а не метода объекта
                sanitized_col_name = sanitize_column_name(column_name)
                if sanitized_col_name.lower() == 'id':
                     sanitized_col_name = f"data_{sanitized_col_name}"

                # row_index в Python (0-based) -> id в БД (1-based)
                db_row_id = row_index + 1

                cursor.execute(f'SELECT "{sanitized_col_name}" FROM {editable_table_name} WHERE id = ?', (db_row_id,))
                old_value_row = cursor.fetchone()
                old_value = old_value_row[0] if old_value_row else None
                logger.debug(f"AppController: Старое значение ячейки: '{old_value}'.")

                # 3. Обновить значение ячейки в таблице editable_data_...
                # ИСПРАВЛЕНО: сигнатура и логика вызова
                update_success = storage.update_editable_cell(sheet_name, row_index, column_name, new_value)

                if update_success:
                    # 4. Сохранить запись в истории редактирования
                    # ИСПРАВЛЕНО: сигнатура вызова, используем правильные аргументы
                    # Старая сигнатура: save_edit_history_record(self, sheet_id: int, operation_type: str, row_index: int, column_name: str, old_value: Any, new_value: Any)
                    # Новая сигнатура из history.py: save_edit_history_record(connection, project_id, sheet_id, cell_address, action_type, old_value, new_value, user, details)
                    
                    # Нужно получить project_id
                    # Предположим, что project_id хранится в self.current_project или можно получить из БД
                    # Получим project_id из таблицы sheets
                    cursor.execute("SELECT project_id FROM sheets WHERE id = ?", (sheet_id,))
                    project_row = cursor.fetchone()
                    if not project_row:
                         logger.error(f"Не удалось получить project_id для sheet_id={sheet_id}.")
                         # Решение: можно считать это критической ошибкой или попытаться получить иначе
                         # Пока считаем критической
                         return False 
                    project_id = project_row[0]
                    
                    # Формируем адрес ячейки, например "A1"
                    # Для этого нужно знать, как сопоставляются column_name и буквы столбцов.
                    # Это может быть сложной логикой. Пока передадим column_name как есть или row_index, column_name.
                    # history.py поддерживает cell_address как опциональный аргумент, так что можно передать None
                    # или сформировать строку вида "R{row_index}C{column_name}" или "{column_name}{row_index+1}"
                    # Для простоты передадим None, а детали положим в `details`
                    cell_address = None # Или сформировать, если нужно
                    
                    history_details = {
                        "row_index": row_index,
                        "column_name": column_name
                    }

                    history_success = storage.save_edit_history_record(
                        project_id=project_id,         # <-- Добавлено
                        sheet_id=sheet_id,
                        cell_address=cell_address,    # <-- Может быть None
                        action_type="edit_cell",      # <-- Исправлен тип действия
                        old_value=old_value,
                        new_value=new_value,
                        user=None,                    # <-- Пока нет пользователей
                        details=history_details       # <-- Передаем детали
                    )

                    if history_success:
                        logger.info(f"Ячейка [{sheet_name}][{row_index}, {column_name}] обновлена и запись истории сохранена.")
                        return True
                    else:
                        logger.error(f"Ячейка обновлена, но ошибка при сохранении записи истории для [{sheet_name}][{row_index}, {column_name}].")
                        # Решение: считать операцию неудачной, если не удалось записать историю?
                        # Или всё же считать успехом обновление данных?
                        # Пока считаем частичный успех как общий успех, но логируем ошибку.
                        # ЛУЧШЕ: если история критична, то возвращаем False.
                        # Для MVP, скорее всего, лучше вернуть True, но залогировать ошибку.
                        # Уточним: если save_edit_history_record возвращает False, это значит, 
                        # что запись в БД не удалась. Это серьезная проблема.
                        # Возвращаем False.
                        return False 
                else:
                    logger.error(f"Не удалось обновить ячейку [{sheet_name}][{row_index}, {column_name}] в БД.")
                    return False
        except Exception as e:
            logger.error(f"Ошибка при обновлении ячейки [{sheet_name}][{row_index}, {column_name}]: {e}", exc_info=True)
            return False

    # ===============================================================

    def get_project_info(self) -> Optional[Dict[str, Any]]:
        """
        Получение информации о текущем проекте.

        Returns:
            Optional[Dict[str, Any]]: Информация о проекте или None если проект не загружен
        """
        if not self.is_project_loaded or not self.current_project:
            logger.warning("Проект не загружен")
            return None
        return self.current_project

    def shutdown(self) -> None:
        """
        Корректное завершение работы приложения.
        """
        logger.info("Завершение работы приложения")

        # Здесь можно добавить сохранение состояния, закрытие соединений и т.д.
        if self.project_manager:
            self.project_manager.cleanup()

        logger.info("Приложение завершено")


# Функции для удобного использования контроллера

def create_app_controller(project_path: Optional[str] = None) -> AppController:
    """
    Фабричная функция для создания экземпляра AppController.

    Args:
        project_path (Optional[str]): Путь к директории проекта

    Returns:
        AppController: Экземпляр контроллера приложения
    """
    controller = AppController(project_path)
    return controller


# Пример использования

if __name__ == "__main__":
    # Это просто для демонстрации, не будет выполняться при импорте
    logger.info("Демонстрация работы AppController")

    # Создание контроллера
    app = create_app_controller()

    # Инициализация
    if app.initialize():
        logger.info("Контроллер приложения инициализирован успешно")
    else:
        logger.error("Ошибка инициализации контроллера приложения")
