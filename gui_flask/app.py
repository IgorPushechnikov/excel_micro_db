# gui_flask/app.py

import os
import tempfile
import uuid
from pathlib import Path
from flask import Flask, render_template, request, jsonify

# Добавим путь к backend в sys.path, чтобы импортировать AppController
import sys
backend_path = Path(__file__).parent.parent / "backend"
sys.path.insert(0, str(backend_path))

# Импортируем AppController
from core.app_controller import create_app_controller


# Создание экземпляра Flask-приложения
# Указываем, что шаблоны находятся в папке ../gui_flask/templates относительно этого файла
app = Flask(__name__, template_folder='../gui_flask/templates', static_folder='../gui_flask/static')

# Создаем временную директорию для проектов внутри gui_flask
temp_projects_dir = Path(__file__).parent / "temp_projects"
temp_projects_dir.mkdir(exist_ok=True)

# Храним активный проект в памяти (в реальном приложении лучше использовать сессии или БД)
# Для MVP этого достаточно
active_project = None
active_project_path = None

@app.route('/')
def index():
    """
    Маршрут для главной страницы GUI.
    """
    # Отображаем базовый шаблон index.html
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Загружает XLSX-файл во временный проект."""
    global active_project, active_project_path
    
    if 'file' not in request.files:
        return jsonify({'error': 'Файл не найден в запросе'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Имя файла пустое'}), 400
    
    if file and file.filename.endswith('.xlsx'):
        try:
            # Создаем уникальную временную директорию для проекта
            project_uuid = str(uuid.uuid4())
            project_path = temp_projects_dir / project_uuid
            project_path.mkdir(exist_ok=True)
            
            # Сохраняем файл
            uploaded_file_path = project_path / file.filename
            file.save(str(uploaded_file_path))
            
            # Инициализируем проект через AppController
            app_controller = create_app_controller(str(project_path))
            if not app_controller.initialize():
                return jsonify({'error': 'Не удалось инициализировать проект'}), 500
            
            # Создаем структуру проекта
            if not app_controller.create_project(str(project_path)):
                return jsonify({'error': 'Не удалось создать структуру проекта'}), 500
            
            # Сохраняем ссылки на активный проект
            active_project = app_controller
            active_project_path = project_path
            
            return jsonify({'message': 'Файл загружен и проект создан', 'project_id': project_uuid, 'filename': file.filename}), 200
        except Exception as e:
            # Очищаем, если что-то пошло не так
            if project_path and project_path.exists():
                import shutil
                shutil.rmtree(project_path, ignore_errors=True)
            return jsonify({'error': f'Ошибка при обработке файла: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Неподдерживаемый тип файла. Загрузите .xlsx файл.'}), 400

@app.route('/analyze', methods=['POST'])
def analyze_file():
    """Анализирует загруженный XLSX-файл."""
    global active_project, active_project_path
    
    if not active_project or not active_project_path:
        return jsonify({'error': 'Нет активного проекта. Загрузите файл сначала.'}), 400
    
    try:
        # Находим загруженный XLSX файл
        xlsx_files = list(active_project_path.glob("*.xlsx"))
        if not xlsx_files:
            return jsonify({'error': 'XLSX файл не найден в проекте'}), 400
        
        xlsx_file_path = xlsx_files[0] # Берем первый найденный
        
        # Выполняем анализ
        success = active_project.analyze_excel_file(str(xlsx_file_path))
        
        if success:
            return jsonify({'message': 'Анализ завершен успешно'}), 200
        else:
            return jsonify({'error': 'Ошибка при анализе файла'}), 500
            
    except Exception as e:
        return jsonify({'error': f'Ошибка при анализе: {str(e)}'}), 500

@app.route('/sheets', methods=['GET'])
def get_sheets():
    """Получает список листов из активного проекта."""
    global active_project
    
    if not active_project or not active_project.storage:
        return jsonify({'error': 'Нет активного проекта или БД не загружена'}), 400
    
    try:
        # Загружаем метаданные всех листов
        sheets_metadata = active_project.storage.load_all_sheets_metadata()
        sheet_names = [sheet['name'] for sheet in sheets_metadata]
        return jsonify({'sheets': sheet_names}), 200
    except Exception as e:
        return jsonify({'error': f'Ошибка при получении списка листов: {str(e)}'}), 500

@app.route('/sheet_data/<sheet_name>', methods=['GET'])
def get_sheet_data(sheet_name):
    """Получает данные листа для отображения в ag-Grid."""
    global active_project
    
    if not active_project or not active_project.storage:
        return jsonify({'error': 'Нет активного проекта или БД не загружена'}), 400
    
    try:
        # Получаем редактируемые данные через DataManager
        # editable_data_dict = active_project.data_manager.get_sheet_editable_data(sheet_name)
        # if not editable_data_dict:
        #     return jsonify({'error': 'Не удалось получить данные листа'}), 500
        
        # Преобразуем данные в формат ag-Grid
        # column_names = editable_data_dict.get('column_names', [])
        # rows_data = editable_data_dict.get('rows', [])
        
        # Для MVP используем более простой подход: загружаем список словарей cell_address -> value
        # и преобразуем его в строки/столбцы
        sheet_id = active_project.data_manager._get_sheet_id_by_name(sheet_name)
        if sheet_id is None:
            return jsonify({'error': f'Лист "{sheet_name}" не найден'}), 404
            
        editable_data_list = active_project.storage.load_sheet_editable_data(sheet_id, sheet_name)
        
        # Преобразование в формат ag-Grid (список объектов {col1: val1, col2: val2, ...})
        # Это упрощенная версия, предполагающая, что данные плотные и начинаются с A1
        grid_data = []
        if editable_data_list:
            # Найдем максимальный номер строки и столбца
            max_row = 0
            max_col = 0
            cell_dict = {}
            for item in editable_data_list:
                addr = item["cell_address"]
                # Простой парсер адреса (только для A1, B2 и т.д., без диапазонов)
                col_part = ""
                row_part = ""
                for char in addr:
                    if char.isalpha():
                        col_part += char
                    else:
                        row_part += char
                if row_part and col_part:
                    row_idx = int(row_part) - 1  # 0-based
                    col_idx = active_project.data_manager._column_letter_to_index(col_part)
                    max_row = max(max_row, row_idx)
                    max_col = max(max_col, col_idx)
                    cell_dict[(row_idx, col_idx)] = item["value"]
            
            # Создаем строки данных для ag-Grid
            for r in range(max_row + 1):
                grid_row = {}
                for c in range(max_col + 1):
                    # Генерируем имя столбца (A, B, ..., Z, AA, AB, ...)
                    col_name = _index_to_column_letter(c)
                    grid_row[col_name] = cell_dict.get((r, c), "")
                grid_data.append(grid_row)
        
        return jsonify(grid_data), 200
        
    except Exception as e:
        import traceback
        traceback.print_exc() # Для отладки
        return jsonify({'error': f'Ошибка при получении данных листа: {str(e)}'}), 500

def _index_to_column_letter(index: int) -> str:
    """Преобразует 0-базовый индекс в имя столбца Excel (A, B, ..., Z, AA, ...)."""
    result = ""
    while index >= 0:
        result = chr(index % 26 + ord('A')) + result
        index = index // 26 - 1
    return result if result else "A"

# Проверка, что скрипт запущен напрямую, а не импортирован
if __name__ == '__main__':
    # Запуск Flask-приложения в режиме разработки
    # host='0.0.0.0' позволяет получить доступ с других устройств в сети (опционально)
    # port=5000 стандартный для Flask
    app.run(debug=True, host='0.0.0.0', port=5000)
