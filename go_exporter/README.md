# Go Excel Exporter

Этот модуль представляет собой автономную утилиту на языке Go, предназначенную для генерации файлов Excel (`.xlsx`) на основе структурированных данных, полученных из основного приложения **Excel Micro DB**.

## Цель

Заменить текущий экспорт на основе `xlsxwriter` (Python) на более мощный и совместимый экспорт с использованием библиотеки `Excelize` (Go). Основное преимущество — полная поддержка редактируемых диаграмм и более широкий спектр типов диаграмм.

## Архитектура

Утилита спроектирована как **CLI-инструмент**, который вызывается из основного Python-приложения. Она не взаимодействует с базой данных напрямую, а получает все необходимые данные через промежуточный JSON-файл.

Такой подход обеспечивает:
1.  **Гибкость:** В будущем основное приложение сможет использовать любую СУБД (SQLite, PostgreSQL, MySQL и т.д.), а Go-экспортер останется неизменным, так как он работает только с JSON.
2.  **Изоляцию:** Go-модуль полностью независим и может быть легко протестирован отдельно.
3.  **Простоту распространения:** Компилируется в один бинарный файл.

## Формат данных (Контракт)

Основное приложение (Python) должно подготовить файл в формате `ExportData.json` со следующей структурой:

```json
{
  "metadata": {
    "project_name": "string",
    "author": "string",
    "created_at": "string (ISO 8601)"
  },
  "sheets": [
    {
      "name": "string",
      "data": [
        ["cell_value_or_null", "cell_value_or_null", ...],
        ...
      ],
      "formulas": [
        {
          "cell": "A1",
          "formula": "=SUM(B1:C1)"
        },
        ...
      ],
      "styles": [
        {
          "range": "A1:C10",
          "style": {
            "font": { "bold": true, "color": "FF0000" },
            "fill": { "fg_color": "FFFF00" },
            "alignment": { "horizontal": "center" }
          }
        },
        ...
      ],
      "charts": [
        {
          "type": "col",
          "position": "E5",
          "title": "My Chart",
          "series": [
            {
              "name": "Sheet1!$A$1",
              "categories": "Sheet1!$B$1:$D$1",
              "values": "Sheet1!$B$2:$D$2"
            }
          ]
        },
        ...
      ]
    }
  ]
}
```

## Использование

После компиляции утилита (`go_excel_exporter.exe` на Windows) вызывается с двумя аргументами:

```bash
./go_excel_exporter.exe path/to/export_data.json path/to/output_file.xlsx
```

### Пример вызова из Python

```python
import subprocess

result = subprocess.run([
    "go_excel_exporter.exe",
    "temp/export_data.json",
    "output/report.xlsx"
], capture_output=True, text=True, check=True)
```

## Сборка

1. Установите Go (версия 1.21+).
2. Перейдите в директорию `go_exporter`.
3. Выполните команду:
   ```bash
   go build -o go_excel_exporter.exe .
   ```

## Зависимости

Основная зависимость — библиотека [`github.com/xuri/excelize/v2`](https://github.com/xuri/excelize).