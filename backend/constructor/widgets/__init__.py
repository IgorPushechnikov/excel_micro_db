# backend/constructor/widgets/__init__.py
"""
Подпакет виджетов графического интерфейса (Constructor) для Excel Micro DB.
"""

# Определяем, что экспортируется при 'from backend.constructor.widgets import *'.
# Пользователи должны импортировать конкретные модули напрямую,
# например, from backend.constructor.widgets.project_explorer import ProjectExplorer.
__all__ = ['ProjectExplorer', 'SheetEditor']

# Импорты удалены из __init__.py для избежания циклических зависимостей.
# При необходимости импортируйте напрямую:
# from .project_explorer import ProjectExplorer
# from .sheet_editor import SheetEditor
