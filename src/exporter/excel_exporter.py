# export_project_single_pass_openpyxl.py
"""Экспорт проекта Excel в один проход с использованием openpyxl."""

import sys
from pathlib import Path
from typing import Dict, Any, List, Optional, Set
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.styles import (
    Font, Fill, Border, PatternFill, Side, Alignment, Protection, NamedStyle, Color
)
from openpyxl.chart import BarChart, PieChart, Reference # Импортируйте нужные типы диаграмм
# from openpyxl.chart._chart import ChartBase # Если нужен общий базовый класс

# --- Импорты или копирование функций создания стилей ---
# Предполагается, что функции _create_openpyxl_*_from_attrs доступны
# Либо импортируем их, либо копируем сюда
# from src.exporter.style_exporter import (
#     _create_openpyxl_font_from_attrs,
#     _create_openpyxl_fill_from_attrs,
#     _create_openpyxl_side_from_attrs,
#     _create_openpyxl_border_from_attrs,
#     _create_openpyxl_alignment_from_attrs,
#     _create_openpyxl_protection_from_attrs,
#     _create_named_style_from_style_attrs # Нужно будет адаптировать под один словарь атрибутов
# )

# --- Копирование функций создания стилей (если не импортируем) ---
# (Здесь должны быть копии функций _create_openpyxl_*_from_attrs и _create_named_style_from_style_attrs)
# Для краткости они опущены, но в рабочем скрипте они должны быть.

# Пример упрощенной функции создания стиля (нужно адаптировать под вашу структуру данных)
def _create_named_style_from_combined_attrs(style_attrs: Dict[str, Any], style_name: str) -> Optional[NamedStyle]:
    """Создает именованный стиль openpyxl из комбинированного словаря атрибутов."""
    # ... (логика из вашего _create_named_style_from_style_attrs, адаптированная под один словарь)
    # Извлекаем подмножества атрибутов для каждого компонента
    # font_attrs = {k.split('_', 1)[1]: v for k, v in style_attrs.items() if k.startswith('font_')}
    # и т.д.
    # Создаем компоненты и добавляем их в NamedStyle
    # named_style = NamedStyle(name=style_name)
    # named_style.font = _create_openpyxl_font_from_attrs(font_attrs)
    # ...
    # return named_style
    pass # Заменить на реальную реализацию

# --- Логика экспорта ---

def export_sheet_data(ws: Worksheet, sheet_data: Dict[str, Any]) -> bool:
    """Экспортирует данные и формулы на лист."""
    try:
        # 1. Создание структуры (заголовки)
        structure = sheet_data.get("structure", [])
        for col_idx, col_info in enumerate(structure, start=1):
            header = col_info.get("column_name", f"Col{col_idx}")
            ws.cell(row=1, column=col_idx, value=header)

        # 2. Заполнение данными
        raw_data = sheet_data.get("raw_data", [])
        for row_idx, row_data in enumerate(raw_data, start=2): # Начинаем со второй строки
             for col_idx, cell_value in enumerate(row_data, start=1):
                 cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                 # Обработка формул, если они есть в cell_value (например, начинаются с '=')
                 # if isinstance(cell_value, str) and cell_value.startswith('='):
                 #     cell.value = cell_value # openpyxl автоматически обработает формулу
        return True
    except Exception as e:
        print(f"Ошибка экспорта данных для листа {ws.title}: {e}")
        return False

def export_sheet_styles(wb: Workbook, ws: Worksheet, styled_ranges_data: List[Dict[str, Any]]) -> bool:
    """Экспортирует стили на лист."""
    try:
        existing_style_names: Set[str] = set(wb.named_styles)
        applied_styles_count = 0

        for style_info in styled_ranges_data:
            range_addr = style_info.get("range_address", "")
            if not range_addr:
                continue

            # Создаем уникальное имя для стиля
            # Адаптируйте под вашу структуру: если атрибуты в подсловаре 'style_attributes'
            # attrs_for_hash = style_info.get("style_attributes", style_info)
            attrs_for_hash = {k: v for k, v in style_info.items() if k != "range_address"}
            style_name = f"Style_{abs(hash(str(sorted(attrs_for_hash.items()))) % 10000000)}"

            # Проверяем и добавляем стиль
            if style_name not in existing_style_names:
                # named_style = _create_named_style_from_style_attrs(attrs_for_hash, style_name) # Оригинальная функция
                named_style = _create_named_style_from_combined_attrs(attrs_for_hash, style_name) # Адаптированная функция
                if named_style:
                    try:
                        wb.add_named_style(named_style)
                        existing_style_names.add(style_name)
                    except Exception as add_style_e:
                         print(f"Предупреждение: Ошибка добавления стиля '{style_name}': {add_style_e}. Продолжаем.")
                         # Проверим, не добавился ли он
                         if style_name not in set(wb.named_styles):
                             print(f"Ошибка: Стиль '{style_name}' не добавлен. Пропущен.")
                             continue
                else:
                    print(f"Ошибка: Не удалось создать именованный стиль для {attrs_for_hash}")
                    continue

            # Применяем стиль к диапазону
            try:
                cell_range = ws[range_addr]
                cells_to_style: List[Cell] = []
                if isinstance(cell_range, Cell):
                    cells_to_style = [cell_range]
                elif hasattr(cell_range, '__iter__'):
                    for item in cell_range:
                        if isinstance(item, Cell):
                            cells_to_style.append(item)
                        elif hasattr(item, '__iter__'):
                            for cell in item:
                                cells_to_style.append(cell)

                for cell in cells_to_style:
                    if style_name in wb._named_styles: # Более эффективная проверка
                        try:
                            cell.style = style_name # КЛЮЧЕВОЙ МОМЕНТ
                        except Exception as apply_e:
                            print(f"Ошибка применения стиля '{style_name}' к {cell.coordinate}: {apply_e}")
                    else:
                         print(f"Предупреждение: Стиль '{style_name}' не найден в книге при применении к {cell.coordinate}.")

                applied_styles_count += 1
            except Exception as apply_range_e:
                print(f"Ошибка применения стиля '{style_name}' к диапазону '{range_addr}': {apply_range_e}")

        print(f"Стили для листа '{ws.title}' применены. Обработано {applied_styles_count} записей.")
        return True
    except Exception as e:
        print(f"Ошибка экспорта стилей для листа {ws.title}: {e}")
        return False

def export_sheet_charts(ws: Worksheet, charts_data: List[Dict[str, Any]]) -> bool:
    """Экспортирует диаграммы на лист."""
    try:
        for chart_info in charts_data:
            # 1. Создание объекта диаграммы
            chart_type = chart_info.get("type", "bar") # Пример: "bar", "pie"
            if chart_type == "bar":
                chart = BarChart()
            elif chart_type == "pie":
                chart = PieChart()
            else:
                print(f"Неизвестный тип диаграммы: {chart_type}. Пропущена.")
                continue

            # 2. Настройка свойств диаграммы
            chart.title = chart_info.get("title", "Chart")
            # chart.style = chart_info.get("style", 2) # Стиль диаграммы
            # chart.x_axis.title = chart_info.get("x_axis_title", "")
            # chart.y_axis.title = chart_info.get("y_axis_title", "")

            # 3. Добавление данных
            # Предполагаем, что данные передаются в виде адресов диапазонов
            data_ref_str = chart_info.get("data_ref", "")
            cats_ref_str = chart_info.get("categories_ref", "")
            if data_ref_str:
                 # Предполагаем, что data_ref_str это строка вида "Sheet!$A$1:$B$10"
                 data_sheet_name, data_range = data_ref_str.split('!', 1)
                 data_sheet = ws.parent[data_sheet_name] if data_sheet_name in ws.parent.sheetnames else ws
                 data = Reference(data_sheet, range_string=data_range)
                 chart.add_data(data, titles_from_data=chart_info.get("titles_from_data", False))

            if cats_ref_str:
                 cats_sheet_name, cats_range = cats_ref_str.split('!', 1)
                 cats_sheet = ws.parent[cats_sheet_name] if cats_sheet_name in ws.parent.sheetnames else ws
                 categories = Reference(cats_sheet, range_string=cats_range)
                 chart.set_categories(categories)

            # 4. Добавление диаграммы на лист
            anchor_cell = chart_info.get("anchor", "A1") # Ячейка для привязки
            ws.add_chart(chart, anchor_cell)

        print(f"Диаграммы для листа '{ws.title}' добавлены.")
        return True
    except Exception as e:
        print(f"Ошибка экспорта диаграмм для листа {ws.title}: {e}")
        return False

def export_project_to_excel_openpyxl(project_data: Dict[str, Any], output_path: str) -> bool:
    """
    Экспортирует весь проект в один файл Excel с использованием openpyxl.
    Args:
        project_data (Dict[str, Any]): Данные проекта.
        output_path (str): Путь для сохранения файла.
    Returns:
        bool: True если успешно, False в противном случае.
    """
    try:
        print("--- Начало экспорта проекта (openpyxl, один проход) ---")
        wb = Workbook()
        # Удаляем дефолтный лист, если он есть
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        sheets_info = project_data.get("sheets", {})

        # --- Этап 1: Данные и формулы ---
        print("Этап 1: Создание структуры и заполнение данными/формулами...")
        for sheet_name, sheet_data in sheets_info.items():
            if sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(title=sheet_name)
            else:
                ws = wb[sheet_name]

            if not export_sheet_data(ws, sheet_data):
                print(f"Ошибка при экспорте данных для листа '{sheet_name}'. Продолжаем.")

        # --- Этап 2: Стили ---
        print("Этап 2: Применение стилей...")
        for sheet_name, sheet_data in sheets_info.items():
             if sheet_name in wb.sheetnames:
                 ws = wb[sheet_name]
                 styled_ranges_data = sheet_data.get("styled_ranges_data", [])
                 if not export_sheet_styles(wb, ws, styled_ranges_data):
                     print(f"Ошибка при экспорте стилей для листа '{sheet_name}'. Продолжаем.")
             else:
                 print(f"Предупреждение: Лист '{sheet_name}' не найден в книге для применения стилей.")

        # --- Этап 3: Диаграммы ---
        print("Этап 3: Добавление диаграмм...")
        for sheet_name, sheet_data in sheets_info.items():
             if sheet_name in wb.sheetnames:
                 ws = wb[sheet_name]
                 charts_data = sheet_data.get("charts_data", [])
                 if not export_sheet_charts(ws, charts_data):
                     print(f"Ошибка при экспорте диаграмм для листа '{sheet_name}'. Продолжаем.")
             else:
                 print(f"Предупреждение: Лист '{sheet_name}' не найден в книге для добавления диаграмм.")

        # --- Сохранение ---
        print(f"Сохранение файла в {output_path}...")
        wb.save(output_path)
        print("--- Экспорт проекта завершен успешно (openpyxl) ---")
        return True

    except Exception as e:
        print(f"Критическая ошибка при экспорте проекта (openpyxl): {e}")
        return False

# --- Пример использования ---
if __name__ == "__main__":
    # project_data = {...} # Загрузите ваши данные проекта
    # output_file = "exported_project_openpyxl.xlsx"
    # export_project_to_excel_openpyxl(project_data, output_file)
    pass
