# build.py
"""
Скрипт для автоматической сборки портабельной версии приложения Excel Micro DB с помощью PyInstaller.

Использование:
    python build.py

Этот скрипт:
1. Проверяет наличие virtual environment.
2. Устанавливает PyInstaller (если нужно).
3. Запускает PyInstaller с настройками из excel_micro_db.spec.
4. Создаёт папку dist/ с портабельной версией.
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.resolve()
DIST_DIR = PROJECT_ROOT / "dist"
SPEC_FILE = PROJECT_ROOT / "excel_micro_db.spec"
ENTRY_POINT = "run_new_gui.py"


def check_venv():
    """Проверяет, запущен ли скрипт внутри виртуального окружения."""
    if sys.prefix == sys.base_prefix:
        print("[!] Предупреждение: Скрипт запущен вне виртуального окружения.")
        print("[i] Рекомендуется активировать venv перед запуском сборки.")
        # Можно продолжить, но предупредить
    else:
        print("[+] Виртуальное окружение активно.")


def install_pyinstaller():
    """Устанавливает PyInstaller, если он не найден."""
    try:
        import PyInstaller
        print("[+] PyInstaller уже установлен.")
    except ImportError:
        print("[i] Установка PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("[+] PyInstaller установлен.")


def run_pyinstaller():
    """
    Запускает PyInstaller с указанным .spec файлом.
    """
    if not SPEC_FILE.exists():
        print(f"[!] Файл спецификации не найден: {SPEC_FILE}")
        sys.exit(1)

    print(f"[i] Запуск PyInstaller с {SPEC_FILE}...")
    try:
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller", str(SPEC_FILE)
        ])
        print("[+] Сборка завершена успешно.")
    except subprocess.CalledProcessError as e:
        print(f"[!] Ошибка при запуске PyInstaller: {e}")
        sys.exit(1)


def main():
    """Основная функция скрипта сборки."""
    print("=== Начало сборки портабельной версии Excel Micro DB ===")

    os.chdir(PROJECT_ROOT)

    check_venv()
    install_pyinstaller()

    # Проверяем точку входа
    if not (PROJECT_ROOT / ENTRY_POINT).exists():
        print(f"[!] Точка входа не найдена: {ENTRY_POINT}")
        sys.exit(1)

    # Очищаем предыдущую сборку
    if DIST_DIR.exists():
        print(f"[i] Очистка предыдущей сборки в {DIST_DIR}...")
        shutil.rmtree(DIST_DIR)

    run_pyinstaller()

    print("\n=== Сборка завершена ===")
    print(f"[i] Портабельная версия находится в: {DIST_DIR}")


if __name__ == "__main__":
    main()
