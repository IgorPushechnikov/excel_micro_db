#!/bin/bash
# setup_env.sh
# Скрипт настройки среды разработки для Excel Micro DB (Linux/Mac)

set -e # Выход при ошибке

echo "================================"
echo "Настройка среды разработки Excel Micro DB"
echo "================================"

# Проверка наличия Python
if ! command -v python3 &> /dev/null; then
    echo "ОШИБКА: Python3 не найден в системе!"
    echo "Пожалуйста, установите Python 3.13 или выше."
    exit 1
fi

PYTHON_VERSION=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
MIN_PYTHON_VERSION="3.13"

if [ "$(printf '%s\n' "$MIN_PYTHON_VERSION" "$PYTHON_VERSION" | sort -V | head -n1)" != "$MIN_PYTHON_VERSION" ]; then
    echo "ОШИБКА: Требуется Python $MIN_PYTHON_VERSION или выше. Найден: $PYTHON_VERSION"
    exit 1
fi

echo "Python найден: $(python3 --version)"


echo ""
echo "Создание виртуального окружения..."
python3 -m venv venv
if [ $? -ne 0 ]; then
    echo "ОШИБКА: Не удалось создать виртуальное окружение!"
    exit 1
fi


echo "Активация виртуального окружения..."
source venv/bin/activate
if [ $? -ne 0 ]; then
    echo "ОШИБКА: Не удалось активировать виртуальное окружение!"
    exit 1
fi


echo ""
echo "Обновление pip..."
python -m pip install --upgrade pip
if [ $? -ne 0 ]; then
    echo "ОШИБКА: Не удалось обновить pip!"
    exit 1
fi


echo ""
echo "Установка основных зависимостей..."
pip install -r ../requirements.txt
if [ $? -ne 0 ]; then
    echo "ОШИБКА: Не удалось установить основные зависимости!"
    exit 1
fi


echo ""
echo "Установка зависимостей для разработки..."
pip install -r requirements-dev.txt
if [ $? -ne 0 ]; then
    echo "ОШИБКА: Не удалось установить зависимости для разработки!"
    exit 1
fi


echo ""
echo "Установка pre-commit хуков..."
if command -v pre-commit &> /dev/null; then
    pre-commit install
    if [ $? -ne 0 ]; then
        echo "ПРЕДУПРЕЖДЕНИЕ: Не удалось установить pre-commit хуки."
    else
        echo "Pre-commit хуки установлены успешно."
    fi
else
    echo "ПРЕДУПРЕЖДЕНИЕ: pre-commit не найден."
fi


echo ""
echo "================================"
echo "Среда разработки настроена успешно!"
echo "================================"
echo ""
echo "Для активации среды разработки в будущем используйте:"
echo "  source venv/bin/activate"
echo ""
echo "Для запуска приложения:"
echo "  python main.py --init"
echo ""
echo "Для запуска GUI:"
echo "  python gui.py"
echo ""
echo "Для запуска тестов:"
echo "  pytest"
echo ""
