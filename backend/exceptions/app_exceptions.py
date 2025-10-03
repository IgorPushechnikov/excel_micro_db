# backend/exceptions/app_exceptions.py
"""
Модуль с пользовательскими исключениями для приложения.
"""


class AppBaseError(Exception):
    """Базовый класс для всех исключений приложения."""
    pass


class ProjectError(AppBaseError):
    """Исключение, связанное с операциями над проектом (создание, загрузка, сохранение)."""
    pass


class AnalysisError(AppBaseError):
    """Исключение, возникающее во время анализа Excel-файлов."""
    pass


class ExportError(AppBaseError):
    """Исключение, возникающее во время экспорта данных проекта."""
    pass


class StorageError(AppBaseError):
    """Исключение, связанное с операциями в хранилище (БД)."""
    pass


class ValidationError(AppBaseError):
    """Исключение, возникающее при валидации данных."""
    pass