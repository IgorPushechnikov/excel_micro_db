# src/storage/metadata.py
"""
Модуль для работы с метаданными проекта в хранилище Excel Micro DB.
(Пока содержит только заглушки или общие функции)
"""
import sqlite3
import json
import logging
from typing import Dict, Any, List, Optional

# from src.storage.base import DateTimeEncoder # Если потребуется
# from src.storage.schema import ... # Если потребуются какие-либо константы

logger = logging.getLogger(__name__)

# Пример функции, которая могла бы быть здесь
# def get_project_info(connection: sqlite3.Connection) -> Dict[str, Any]:
#     """Получает общую информацию о проекте."""
#     pass

# def get_sheet_list(connection: sqlite3.Connection) -> List[str]:
#     """Получает список имен листов."""
#     pass

# def load_sheet_structure(connection: sqlite3.Connection, sheet_name: str) -> List[Dict[str, Any]]:
#     """Загружает структуру листа."""
#     pass
