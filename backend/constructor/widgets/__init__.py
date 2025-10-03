# backend/constructor/widgets/__init__.py

# Импорты для удобства
from .project_explorer import ProjectExplorer
# Импорт SheetEditor из его подмодуля
from .sheet_editor.sheet_editor_widget import SheetEditor

# Определение того, что экспортируется при "from . import *"
# Теперь включает классы, которые мы явно импортировали выше.
__all__ = [
    'ProjectExplorer',
    'SheetEditor',
]
