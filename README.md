# Excel Micro DB

**Микро-СУБД с визуальным конструктором для анализа и воссоздания логики Excel файлов.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-green.svg)](https://github.com/IgorPushechnikov/excel_micro_db)

## 🎯 Цель проекта

Excel Micro DB — это инструмент, который позволяет:

- **Анализировать** сложные Excel файлы и извлекать из них логику вычислений.
- **Документировать** структуру данных и формулы в понятном формате.
- **Воссоздавать** логику в виде мини-СУБД с возможностью обработки и экспорта данных.
- **Визуализировать** данные и логику через удобный графический интерфейс.
- **Интегрировать** данные из различных источников (Excel, БД, интернет).

## 🚀 Основные возможности

- **Анализ Excel**: Извлечение структуры данных, формул и зависимостей.
- **Документация**: Автоматическая генерация документации логики.
- **Хранилище**: Встроенная SQLite база данных для хранения данных и метаинформации.
- **Конструктор**: **Новый** визуальный интерфейс (Electron + React + Vite + Tailwind CSS), приближенный к Excel, с поддержкой нодового редактора.
- **Процессор**: Обработка данных по заданным правилам.
- **Экспорт**: Генерация отчетов и диаграмм в различных форматах.
- **Нодовый редактор**: Визуальное создание сложных формул и скриптов (в разработке).

## 🛠 Технический стек

### Backend (Python)
- **Язык**: Python 3.13+ (разработка ведётся на Python 3.13.7)
- **Фреймворки**: FastAPI (для API, если используется), pandas, numpy, openpyxl
- **БД**: SQLite (встроенная)
- **Экспорт**: xlsxwriter
- **Конфигурация**: YAML
- **Тестирование**: pytest

### Frontend (GUI)
- **Язык**: TypeScript
- **Фреймворк**: React 18+
- **Сборка**: Vite
- **Стилизация**: Tailwind CSS
- **Таблица**: ag-Grid Community
- **Нодовый редактор**: React Flow
- **Десктопное окружение**: Electron
- **Менеджер пакетов**: npm

## 📁 Структура проекта

excel_micro_db/
├── backend/           # Основной Python-бэкенд
├── config/            # Конфигурационные файлы
├── data/              # Входные и обработанные данные
├── dev_env/           # Скрипты и зависимости для разработки
├── docs/              # Документация проекта
├── frontend/          # **Новый GUI** на Electron + React + Vite
├── logs/              # Лог-файлы
├── scripts/           # Вспомогательные скрипты
├── templates/         # Шаблоны
├── tests/             # Тесты
├── test_workspace/    # Рабочая область для тестов
├── venv/              # Виртуальное окружение (обычно в .gitignore)
└── main.py            # Основной CLI/HTTP-сервер

## 📖 Документация

Подробная документация находится в директории `docs/`.

## ▶️ Режимы работы

### 1. **CLI режим** (разработка/автоматизация)

```bash
python main.py --init --project-path ./my_project
python main.py --analyze ./data/input.xlsx
python main.py --process --config config/batch.yaml
```

### 2. **GUI режим** (пользовательский интерфейс)

```bash
# Перейти в директорию GUI
cd frontend

# Установить зависимости
npm install

# Запустить GUI в режиме разработки
npm run dev:frontend  # Запускает Vite
npm run start:dev     # Запускает Electron (в отдельном терминале)
```

### 3. **Интерактивный режим** (REPL для разработчиков)

```bash
python -i main.py --interactive
```

### 4. **Скомпилированный режим** (релиз)

```bash
# Сборка GUI с Electron (в разработке)
# npm run build:electron
```

## 🏗️ Быстрый старт (Backend)

### Установка

1. Клонируйте репозиторий:

   ```bash
   git clone https://github.com/IgorPushechnikov/excel_micro_db.git
   cd excel_micro_db
   ```

2. Настройте среду разработки:

   ```bash
   cd dev_env
   setup_env.bat # Windows
   # или
   ./setup_env.sh # Linux/Mac
   ```

3. Активируйте виртуальное окружение:

   ```bash
   # Windows
   venv\Scripts\activate
   # Linux/Mac
   source venv/bin/activate
   ```

4. Установите зависимости:

   ```bash
   pip install -r requirements.txt
   pip install -r dev_env/requirements-dev.txt
   ```

## 🏗️ Быстрый старт (GUI)

1. Убедитесь, что установлен Node.js и npm.
2. Перейдите в директорию `frontend`.
3. Установите зависимости: `npm install`
4. Запустите разработку: `npm run dev:frontend` и `npm run start:dev`

## 🧪 Тестирование

Для запуска тестов Python используйте pytest:

```bash
pytest
```

Также доступен скрипт для интеграционного теста:

```bash
python scripts/run_integration_test.py
```

## 📄 Лицензия

Этот проект лицензирован по лицензии MIT - подробности см. в файле `LICENSE`.

## 👥 Авторы

- **Игорь Пушечников** - *Идея и разработка* - [IgorPushechnikov](https://github.com/IgorPushechnikov)
