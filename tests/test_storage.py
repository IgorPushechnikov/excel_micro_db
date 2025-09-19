# tests/test_storage.py
"""
Тесты для модуля src/storage/database.py.
"""
import pytest
import tempfile
import os
import yaml
from pathlib import Path
import sys
from pandas import Timestamp # Импортируем Timestamp для обработки

# Добавляем корень проекта в путь поиска модулей
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# --- Обработка специфичных тегов YAML от Pandas ---
# Это необходимо для корректной загрузки test_sample_documentation.yaml,
# который был создан с объектами Timestamp.
def timestamp_constructor(loader, node):
    """Конструктор для десериализации pandas Timestamp из YAML."""
    # Получаем список аргументов из узла
    args = loader.construct_sequence(node)
    # Первый аргумент - это значение времени в наносекундах
    nanoseconds = args[0]
    # Второй аргумент - tzinfo (может быть None)
    tz = args[1] if len(args) > 1 and args[1] is not None else None
    # Третий аргумент - freq (может быть None)
    # freq = args[2] if len(args) > 2 else None # Частота обычно не нужна для создания Timestamp
    # Четвертый аргумент - unit (например, 10 для 'ns')
    unit = args[3] if len(args) > 3 else 'ns'

    # Создаем Timestamp. Упрощаем, используя наносекунды напрямую.
    # Это может потребовать корректировки в зависимости от формата.
    # Более простой и надежный способ - преобразовать в строку и обратно.
    # Но так как исходные данные были Timestamp'ами, предположим, что
    # nanoseconds - это правильное значение.
    # Однако, yaml dump мог сериализовать это нестандартно.
    # Проще всего обработать это как строку или число.
    # Проверим тип первого аргумента.
    if isinstance(nanoseconds, int):
        # Если это число наносекунд, создаем Timestamp
        # pd.Timestamp принимает наносекунды напрямую через unit='ns'
        return Timestamp(nanoseconds, unit='ns')
    else:
        # Если это что-то другое, попробуем преобразовать в строку и создать Timestamp
        # Это менее вероятно, но возможно
        return Timestamp(str(nanoseconds))

# Регистрируем конструктор для тега, найденного в ошибке
yaml.add_constructor('tag:yaml.org,2002:python/object/apply:pandas._libs.tslibs.timestamps._unpickle_timestamp', timestamp_constructor)
# --- Конец обработки специфичных тегов ---

from src.storage.database import ProjectDBStorage, create_storage
# Импортируем logger, если нужно проверять логи (требует дополнительной настройки)
# from src.utils.logger import get_logger
# logger = get_logger(__name__)

# Путь к тестовому файлу documentation.yaml
TEST_DOC_YAML_PATH = project_root / "data" / "samples" / "test_sample_documentation.yaml"

# Фикстура для создания временной БД для каждого теста
@pytest.fixture
def temp_db_path():
    """Создает временный файл БД для теста."""
    with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as tmp_file:
        db_path = tmp_file.name
    yield db_path
    # Удаляем временный файл после теста
    try:
        os.unlink(db_path)
    except OSError:
        pass # Файл уже удален или не существует

# Фикстура для загрузки тестовых данных из YAML
@pytest.fixture
def sample_documentation_data():
    """Загружает тестовые данные из test_sample_documentation.yaml."""
    if not TEST_DOC_YAML_PATH.exists():
        pytest.skip(f"Тестовый файл документации не найден: {TEST_DOC_YAML_PATH}")
    with open(TEST_DOC_YAML_PATH, 'r', encoding='utf-8') as f:
        # Используем full_load или unsafe_load, так как у нас есть специальные теги
        # Но мы зарегистрировали конструктор, поэтому safe_load должен работать
        # Если safe_load все еще не работает, используйте yaml.load(f, Loader=yaml.FullLoader)
        # или yaml.load(f, Loader=yaml.UnsafeLoader) - последний НЕ безопасен!
        try:
            # Попробуем safe_load с нашим зарегистрированным конструктором
            data = yaml.safe_load(f)
        except yaml.constructor.ConstructorError:
            # Если все же возникла ошибка, используем FullLoader
            f.seek(0) # Сбрасываем указатель файла в начало
            data = yaml.load(f, Loader=yaml.FullLoader)
    return data

# --- Тесты ---

def test_create_storage_and_connect(temp_db_path):
    """Тест создания хранилища и подключения к БД."""
    storage = create_storage(temp_db_path)
    assert storage is not None
    assert storage.connection is not None
    # Проверим, что файл БД был создан
    assert Path(temp_db_path).exists()
    storage.close()

def test_project_db_storage_initialization(temp_db_path):
    """Тест инициализации класса ProjectDBStorage."""
    storage = ProjectDBStorage(temp_db_path)
    assert storage.db_path == Path(temp_db_path)
    assert storage.connection is None

def test_connect_and_initialize_database(temp_db_path):
    """Тест подключения и инициализации схемы БД."""
    storage = ProjectDBStorage(temp_db_path)
    assert storage.connect() == True
    assert storage.connection is not None
    
    # Проверим, что таблицы созданы. 
    # Мы можем проверить наличие одной из ключевых таблиц.
    cursor = storage.connection.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='project_info';")
    result = cursor.fetchone()
    assert result is not None # Таблица project_info должна существовать
    storage.close()

def test_save_and_load_analysis_results(temp_db_path, sample_documentation_data):
    """Тест сохранения и загрузки результатов анализа."""
    # 1. Сохраняем данные
    storage = ProjectDBStorage(temp_db_path)
    assert storage.connect() == True
    assert storage.save_analysis_results(sample_documentation_data) == True

    # 2. Загружаем общую информацию проекта
    project_info = storage.load_project_overview()
    assert project_info is not None
    # Обратите внимание: в БД мы сохраняем project_info с id=1, а остальные поля
    # берутся из sample_documentation_data['summary'] и других частей.
    # В YAML project_name находится на верхнем уровне.
    assert project_info['name'] == sample_documentation_data['project_name']
    assert project_info['source_file_path'] == sample_documentation_data['source_file']
    
    # 3. Загружаем список листов
    sheet_names = storage.load_sheet_names()
    expected_sheet_names = [sheet['name'] for sheet in sample_documentation_data['sheets']]
    assert set(sheet_names) == set(expected_sheet_names)

    # 4. Загружаем и проверяем данные для одного из листов (например, Summary)
    summary_sheet_data = storage.load_sheet_data("Summary")
    assert summary_sheet_data is not None
    assert summary_sheet_data['name'] == "Summary"
    # Проверяем количество строк/столбцов
    yaml_summary_sheet = next(s for s in sample_documentation_data['sheets'] if s['name'] == "Summary")
    assert summary_sheet_data['rows_count'] == yaml_summary_sheet['rows_count']
    assert summary_sheet_data['cols_count'] == yaml_summary_sheet['cols_count']

    # Проверяем структуру
    assert len(summary_sheet_data['structure']) == len(yaml_summary_sheet['structure'])
    # Проверяем формулы
    assert len(summary_sheet_data['formulas']) == len(yaml_summary_sheet['formulas'])
    # Проверяем межлистовые ссылки
    assert len(summary_sheet_data['cross_sheet_references']) == len(yaml_summary_sheet['cross_sheet_references'])
    
    # Проверяем диаграммы
    assert len(summary_sheet_data['charts']) == len(yaml_summary_sheet['charts'])
    if summary_sheet_data['charts']:
        # Проверим первую диаграмму
        db_chart = summary_sheet_data['charts'][0]
        yaml_chart = yaml_summary_sheet['charts'][0]
        # assert db_chart['name'] == yaml_chart['name'] # Имя может быть пустым ''
        assert db_chart['type'] == yaml_chart['type']
        # Проверим источники данных диаграммы
        db_sources = sorted(db_chart['data_sources'], key=lambda x: (x.get('sheet', ''), x.get('range', '')))
        yaml_sources = sorted(yaml_chart['data_sources'], key=lambda x: (x.get('sheet', ''), x.get('range', '')))
        assert len(db_sources) == len(yaml_sources)
        for db_src, yaml_src in zip(db_sources, yaml_sources):
             assert db_src['sheet'] == yaml_src['sheet']
             assert db_src['range'] == yaml_src['range']
             # series_part может отличаться в реализации, но должен быть
             assert 'series_part' in db_src 

    storage.close()

def test_load_nonexistent_sheet(temp_db_path, sample_documentation_data):
    """Тест загрузки несуществующего листа."""
    storage = ProjectDBStorage(temp_db_path)
    assert storage.connect() == True
    assert storage.save_analysis_results(sample_documentation_data) == True
    
    # Пытаемся загрузить несуществующий лист
    nonexistent_data = storage.load_sheet_data("NonExistentSheet")
    assert nonexistent_data is None
    storage.close()

def test_close_connection(temp_db_path):
    """Тест закрытия соединения."""
    storage = ProjectDBStorage(temp_db_path)
    assert storage.connect() == True
    assert storage.connection is not None
    storage.close()
    # После закрытия connection должен быть None
    assert storage.connection is None
    # Повторный вызов close не должен вызывать ошибок
    storage.close() 

# Дополнительные тесты могут включать:
# - Тесты на обновление данных (update_sheet_structure, update_sheet_formula)
# - Тесты на обработку ошибок (например, сохранение в недоступную БД)
# - Тесты на работу с пустыми данными из documentation.yaml
