# backend/constructor/widgets/__init__.py

# Импорты для удобства
# from .project_explorer import ProjectExplorer # УБРАНО
# from .sheet_editor.sheet_editor_widget import SheetEditor # УБРАНО

# Определение того, что экспортируется при "from . import *"
# Оставлены только модули/подмодули, присутствующие в папке.
# Убраны классы, так как их импорт вызывал проблемы.
__all__ = [
    # 'ProjectExplorer', # Убрано из __all__
    # 'SheetEditor', # Убрано из __all__
    # 'sheet_editor', # Можно включить, если нужно импортировать подмодуль
    # 'project_explorer', # project_explorer.py как модуль
]

# Явно импортировать подмодули, если нужно, чтобы 'from widgets import ...' работало.
# from . import sheet_editor # Пока закомментировано.
# from . import project_explorer # Пока закомментировано.

# Оставим __init__.py пустым или с минимальными импортами, как было изначально.
# Это устраняет потенциальные проблемы с циклическими импортами или sys.path.
