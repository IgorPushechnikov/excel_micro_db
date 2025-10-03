# backend/constructor/__init__.py

# Импорты для удобства
from .main_window import MainWindow
from .widgets.project_explorer import ProjectExplorer
# Импорт SheetEditor из его подмодуля
from .widgets.sheet_editor.sheet_editor_widget import SheetEditor

# Определение того, что экспортируется при "from . import *"
# Теперь включает классы, которые мы явно импортировали выше.
__all__ = [
    'MainWindow',
    'ProjectExplorer',
    'SheetEditor',
    # 'gui_app', # gui_app - это модуль, его обычно не включают в __all__, если не планируют использовать как объект
    # Если gui_app нужно экспортировать как модуль, раскомментируйте следующую строку:
    # 'gui_app',
]
