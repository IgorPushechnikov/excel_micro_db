# backend/constructor/__init__.py

# Импорты для удобства
# from .main_window import MainWindow # УБРАНО
# from .widgets.project_explorer import ProjectExplorer # УБРАНО
# from .widgets.sheet_editor.sheet_editor_widget import SheetEditor # УБРАНО

# Определение того, что экспортируется при "from . import *"
# Оставлены только модули, присутствующие в папке как .py файлы.
# Убраны классы, так как их импорт вызывал проблемы.
__all__ = [
    # 'MainWindow', # Убрано из __all__
    # 'ProjectExplorer', # Убрано из __all__
    # 'SheetEditor', # Убрано из __all__
    # 'gui_app', # gui_app - это модуль, его обычно не включают в __all__, если не планируют использовать как объект
    # Если gui_app нужно экспортировать как модуль, раскомментируйте следующую строку:
    # 'gui_app',
    # Если нужно экспортировать подмодули (папки с __init__.py), можно включить их имена.
    # Однако, в текущей структуре, подмодулем является 'widgets', но он может не быть виден,
    # если __init__.py пуст или не содержит from . import widgets
    # Пока оставим __all__ пустым или с явным указанием модулей, если они нужны.
    # В данном случае, файл gui_app.py присутствует.
    'gui_app', # Включаем модуль gui_app, если он используется как точка входа.
    # 'main_window', # Можно добавить, если используется.
    # 'widgets', # widgets - это папка. Для её импорта нужно, чтобы __init__.py в constructor его явно импортировал.
]

# Явно импортируем модуль gui_app, если он используется как точка входа.
# Это делает его доступным при 'from constructor import gui_app'
from . import gui_app

# Если нужно, можно также импортировать подмодуль 'widgets',
# чтобы 'from constructor import widgets' работало.
# from . import widgets # Пока закомментировано.
