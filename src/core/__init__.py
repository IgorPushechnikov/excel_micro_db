# src/core/__init__.py
"""
Инициализация модуля core.
Предоставляет доступ к основным классам, таким как AppController, из старого расположения.
"""

# Импортируем AppController из новой поддиректории для обратной совместимости
try:
    from .controller.app_controller import AppController, create_app_controller
    # Если нужно, можно импортировать и другие классы из controller
    # from .controller.data_manager import DataManager
    # from .controller.project_manager import ProjectManager
except ImportError as e:
    # Логируем ошибку, если модуль не может быть импортирован
    import logging
    logger = logging.getLogger(__name__)
    logger.error(f"Не удалось импортировать AppController из controller: {e}")
    # Важно: Не вызываем raise, чтобы не ломать импорт в целом,
    # но разработчик увидит ошибку в логах.
    # AppController будет недоступен через этот __init__.py, если файл перемещения не найден.

# Также можно импортировать ProjectManager, если он теперь в controller
# или остался в корне core. Проверим его расположение.
# Текущий `ProjectManager` находится в `src/core/project_manager.py`.
# Если он останется там, его можно импортировать здесь.
# from .project_manager import ProjectManager # Пример, если потребуется
