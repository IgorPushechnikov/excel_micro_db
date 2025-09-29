# src/exceptions/app_exceptions.py
"""
Пользовательские исключения для приложения Excel Micro DB.
"""

class ProjectError(Exception):
    """Базовый класс для исключений, связанных с проектом."""
    pass

class AnalysisError(Exception):
    """Исключение, возникающее при ошибке анализа Excel-файла."""
    pass

class ExportError(Exception):
    """Исключение, возникающее при ошибке экспорта данных."""
    pass

# Дополнительные пользовательские исключения можно добавлять здесь
