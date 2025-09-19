# check_db_content.py
import sqlite3
from pathlib import Path

# Укажи путь к БД
db_path = Path("test_workspace/integration_test_project_20250828_163834/project_data.db")

def check_db(db_path: Path):
    if not db_path.exists():
        print(f"[ERROR] БД не найдена: {db_path}")
        return

    print(f"[INFO] Подключение к БД: {db_path}")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # 1. Список таблиц
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [row[0] for row in cursor.fetchall()]
    print("\n[Таблицы в БД]:")
    for t in tables:
        print(f"  - {t}")

    # 2. Информация о проекте
    print("\n[project_info]:")
    cursor.execute("SELECT * FROM project_info;")
    rows = cursor.fetchall()
    if rows:
        for row in rows:
            print(dict(row))
    else:
        print("  (пусто)")

    # 3. Листы
    print("\n[sheets]:")
    cursor.execute("SELECT * FROM sheets;")
    rows = cursor.fetchall()
    if rows:
        for row in rows:
            print(dict(row))
    else:
        print("  (пусто)")

    # 4. Структура листов
    print("\n[sheet_structure]:")
    cursor.execute("SELECT * FROM sheet_structure LIMIT 5;")  # Ограничим для краткости
    rows = cursor.fetchall()
    if rows:
        for row in rows:
            print(dict(row))
    else:
        print("  (пусто)")

    # 5. Формулы
    print("\n[sheet_formulas]:")
    cursor.execute("SELECT * FROM sheet_formulas LIMIT 5;")
    rows = cursor.fetchall()
    if rows:
        for row in rows:
            print(dict(row))
    else:
        print("  (пусто)")

    # 6. Диаграммы
    print("\n[sheet_charts]:")
    cursor.execute("SELECT * FROM sheet_charts;")
    rows = cursor.fetchall()
    if rows:
        for row in rows:
            print(dict(row))
    else:
        print("  (пусто)")

    # 7. Источники данных диаграмм
    print("\n[sheet_chart_data_sources]:")
    cursor.execute("SELECT * FROM sheet_chart_data_sources;")
    rows = cursor.fetchall()
    if rows:
        for row in rows:
            print(dict(row))
    else:
        print("  (пусто)")

    conn.close()
    print(f"\n[INFO] Проверка БД завершена.")

if __name__ == "__main__":
    check_db(db_path)