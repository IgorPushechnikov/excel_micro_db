# backend/constructor/__init__.py
"""
Пакет графического интерфейса (Constructor) для Excel Micro DB.
"""

# Определяем, что экспортируется при 'from backend.constructor import *'.
# Пользователи должны импортировать конкретные модули напрямую,
# например, from backend.constructor.main_window import MainWindow.
__all__ = ['MainWindow', 'ProjectExplorer', 'SheetEditor']

# Импорты удалены из __init__.py для избежания циклических зависимостей.
# При необходимости импортируйте напрямую:
# from .main_window import MainWindow
# from .widgets.project_explorer import ProjectExplorer
# from .widgets.sheet_editor import SheetEditor
