#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для тестирования запуска GUI Excel Micro DB.
"""

import sys
import os
from pathlib import Path

# Добавляем директорию проекта в путь Python
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Импортируем функцию main из gui_app
from src.constructor.gui_app import main as gui_main

if __name__ == "__main__":
    print("Начинаем тестовый запуск GUI...")
    try:
        # Вызываем main функцию из gui_app
        exit_code = gui_main()
        print(f"GUI завершился с кодом: {exit_code}")
    except Exception as e:
        print(f"Произошла ошибка при запуске GUI: {e}")
        import traceback
        traceback.print_exc()
