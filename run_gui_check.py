#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для проверки запуска GUI Excel Micro DB.
Показывает подробные сообщения о статусе и ошибках.
"""

import sys
import os
from pathlib import Path
import traceback

# Добавляем корень проекта в путь поиска модулей
project_root = Path(__file__).resolve().parent
sys.path.insert(0, str(project_root))

print(f"Текущая директория: {Path.cwd()}")
print(f"Корень проекта добавлен в sys.path: {project_root}")

# --- Проверка зависимостей ---
dependencies = [
    'PySide6',
    'numpy',
    'pandas',
    'openpyxl',
    'xlsxwriter',
    'matplotlib',
    'yaml',
    'dotenv'
]

missing_deps = []
for dep in dependencies:
    try:
        __import__(dep)
        print(f"[OK] Зависимость '{dep}' найдена")
    except ImportError:
        print(f"[ОШИБКА] Зависимость '{dep}' НЕ НАЙДЕНА")
        missing_deps.append(dep)

if missing_deps:
    print(f"\n!!! ОШИБКА: Не найдены следующие зависимости: {', '.join(missing_deps)}")
    print("Выполните: pip install -r requirements.txt")
    sys.exit(1)
else:
    print("\nВсе зависимости найдены. Проверяем импорт модулей GUI...\n")

# --- Проверка импорта модулей GUI ---
try:
    from src.constructor.gui_app import main as gui_main
    print("[OK] Успешно импортирован gui_app.main")
except ImportError as e:
    print(f"[ОШИБКА] Не удалось импортировать gui_app.main: {e}")
    traceback.print_exc()
    sys.exit(1)

try:
    from src.constructor.gui_controller import GUIController
    print("[OK] Успешно импортирован GUIController")
except ImportError as e:
    print(f"[ОШИБКА] Не удалось импортировать GUIController: {e}")
    traceback.print_exc()
    sys.exit(1)

try:
    from src.constructor.main_window import MainWindow
    print("[OK] Успешно импортирован MainWindow")
except ImportError as e:
    print(f"[ОШИБКА] Не удалось импортировать MainWindow: {e}")
    traceback.print_exc()
    sys.exit(1)

try:
    from src.core.app_controller import create_app_controller
    print("[OK] Успешно импортирован create_app_controller")
except ImportError as e:
    print(f"[ОШИБКА] Не удалось импортировать create_app_controller: {e}")
    traceback.print_exc()
    sys.exit(1)

print("\nВсе модули GUI успешно импортированы.\n")

# --- Попытка запуска GUI ---
print("Попытка запуска GUI...")
try:
    exit_code = gui_main()
    print(f"GUI завершился с кодом: {exit_code}")
except Exception as e:
    print(f"Критическая ошибка при запуске GUI: {e}")
    traceback.print_exc()
    sys.exit(1)
