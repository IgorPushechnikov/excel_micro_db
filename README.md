# Excel Micro DB

**Лёгкая СУБД для хранения и анализа Excel-файлов.**

## 🧰 Описание

Excel Micro DB — это инструмент для импорта Excel-файлов в локальную SQLite БД с сохранением:

- Значений ячеек
- Формул
- Стилей
- Диаграмм
- Метаданных
- Имён листов
- Объединённых ячеек
- И многого другого

Поддерживает импорт через `openpyxl` и `xlwings`.

## 🚀 Быстрый старт

### 1. Клонирование репозитория

```bash
git clone <URL_репозитория>
cd excel_micro_db
```

### 2. Создание виртуального окружения

```bash
python -m venv venv
```

### 3. Активация виртуального окружения

#### Windows (cmd):

```cmd
venv\Scripts\activate
```

#### Windows (PowerShell):

```powershell
venv\Scripts\Activate.ps1
```

> Если PowerShell ругается:
>
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

#### Linux/macOS:

```bash
source venv/bin/activate
```

### 4. Установка зависимостей

```bash
pip install -r requirements.txt
```

### 5. Запуск GUI

```bash
python run_new_gui.py
```

### 6. Деактивация окружения

```bash
deactivate
```

## 📦 Структура проекта

```
excel_micro_db/
├── backend/           # Основной код приложения
│   ├── core/          # Ядро: AppController
│   ├── storage/       # Работа с БД (SQLite)
│   ├── importer/      # Импорт из Excel (openpyxl, xlwings)
│   ├── constructor/   # GUI (Qt)
│   ├── analyzer/      # Анализ данных
│   ├── exporter/      # Экспорт в Excel
│   └── utils/         # Вспомогательные модули
├── data/              # Примеры данных
├── config/            # Конфигурационные файлы
├── logs/              # Логи
├── tests/            # Тесты
├── run_new_gui.py     # Точка входа в GUI
├── requirements.txt   # Зависимости
└── README.md          # Этот файл
```

## 🧪 Импорт данных

- Через меню **Импорт → Импорт данных...** (через `openpyxl`).
- Через меню **Импорт → Забрать из Excel** (через `xlwings`).

## 🛠️ Сборка exe

```bash
python build.py
```

Результат: `dist/excel_micro_db.exe`.

## 📄 Лицензия

MIT
