@echo off
:: setup_env.bat
:: Скрипт настройки среды разработки для Excel Micro DB (Windows)

echo ================================
echo Настройка среды разработки Excel Micro DB
echo ================================

:: Проверка наличия Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ОШИБКА: Python не найден в системе!
    echo Пожалуйста, установите Python 3.13 или выше.
    pause
    exit /b 1
)

echo Python найден:
python --version

echo.
echo Создание виртуального окружения...
python -m venv venv
if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось создать виртуальное окружение!
    pause
    exit /b 1
)

echo Активация виртуального окружения...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось активировать виртуальное окружение!
    pause
    exit /b 1
)

echo.
echo Обновление pip...
python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось обновить pip!
    pause
    exit /b 1
)

echo.
echo Установка основных зависимостей...
pip install -r ..\requirements.txt
if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось установить основные зависимости!
    pause
    exit /b 1
)

echo.
echo Установка зависимостей для разработки...
pip install -r requirements-dev.txt
if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось установить зависимости для разработки!
    pause
    exit /b 1
)

echo.
echo Установка pre-commit хуков...
pre-commit install
if %errorlevel% neq 0 (
    echo ПРЕДУПРЕЖДЕНИЕ: Не удалось установить pre-commit хуки. Продолжаем...
) else (
    echo Pre-commit хуки установлены успешно.
)

echo.
echo ================================
echo Среда разработки настроена успешно!
echo ================================
echo.
echo Для активации среды разработки в будущем используйте:
echo   call venv\Scripts\activate.bat
echo.
echo Для запуска приложения:
echo   python main.py --init
echo.
echo Для запуска GUI:
echo   python gui.py
echo.
echo Для запуска тестов:
echo   pytest
echo.
pause