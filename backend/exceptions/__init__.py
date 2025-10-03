# backend/exceptions/__init__.py
"""
Пакет пользовательских исключений для приложения.
"""

from .app_exceptions import (
    AppBaseError,
    ProjectError,
    AnalysisError,
    ExportError,
    StorageError,
    ValidationError,
)

__all__ = [
    "AppBaseError",
    "ProjectError",
    "AnalysisError",
    "ExportError",
    "StorageError",
    "ValidationError",
]