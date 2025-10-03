# backend/constructor/widgets/sheet_editor.py
"""
Точка входа для виджета-редактора листа Excel.
Импортирует и экспортирует основные классы из подмодулей.
"""

# Импортируем основные классы из подмодулей
from .sheet_editor.sheet_data_model import SheetDataModel
from .sheet_editor.sheet_editor_widget import SheetEditor, EditAction

# Экспортируем их для удобства использования
__all__ = ['SheetDataModel', 'SheetEditor', 'EditAction']

# Если нужно, можно добавить здесь дополнительную логику инициализации пакета
# Например, проверку зависимостей или регистрацию плагинов (если применимо)
