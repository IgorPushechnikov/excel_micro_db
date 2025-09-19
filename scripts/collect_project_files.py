# scripts/collect_project_files.py
"""
Скрипт для сбора содержимого текстовых файлов проекта в одну папку.
Игнорирует стандартные файлы кэша/сборки Python и двоичные файлы.
Полезно для отправки файлов в чат или анализа.
"""
import os
import shutil
import mimetypes
from pathlib import Path

# Стандартный список игнорируемых путей/файлов (аналогично .gitignore)
DEFAULT_IGNORE_PATTERNS = {
    # Byte-compiled / optimized / DLL files
    '__pycache__/', '*.py[cod]', '*$py.class',
    # C extensions
    '*.so',
    # Distribution / packaging
    '.Python', 'build/', 'develop-eggs/', 'dist/', 'downloads/', 'eggs/',
    '*.egg-info/', '.installed.cfg', '*.egg', 'MANIFEST',
    # PyInstaller
    '*.manifest', '*.spec',
    # Installer logs
    'pip-log.txt', 'pip-delete-this-directory.txt',
    # Unit test / coverage reports
    'htmlcov/', '.tox/', '.nox/', '.coverage', '.coverage.*', '.cache',
    'nosetests.xml', 'coverage.xml', '*.cover', '*.py,cover', '.hypothesis/',
    '.pytest_cache/', 'cover/',
    # Translations
    '*.mo', '*.pot',
    # Django stuff:
    '*.log', 'local_settings.py', 'db.sqlite3', 'db.sqlite3-journal',
    # Flask stuff:
    'instance/', '.webassets-cache',
    # Scrapy stuff:
    '.scrapy',
    # Sphinx documentation
    'docs/_build/',
    # PyBuilder
    '.pybuilder/',
    # Jupyter Notebook
    '.ipynb_checkpoints',
    # IPython
    'profile_default/', 'ipython_config.py',
    # pyenv
    '.python-version',
    # pipenv
    'Pipfile', 'Pipfile.lock',
    # PEP 582
    '__pypackages__/',
    # Celery stuff
    'celerybeat-schedule', 'celerybeat.pid',
    # SageMath parsed files
    '*.sage.py',
    # Environments
    '.env', '.venv', 'venv/', 'ENV/', 'env.bak/', 'venv.bak/',
    # Spyder project settings
    '.spyderproject', '.spyproject',
    # Rope project settings
    '.ropeproject',
    # mkdocs documentation
    '/site',
    # mypy
    '.mypy_cache/', '.dmypy.json', 'dmypy.json',
    # Pyre type checker
    '.pyre/',
    # IDE
    '.vscode/', '.idea/', '*.swp', '*.swo',
    # OS
    '.DS_Store', 'Thumbs.db',
    # SQLite databases (if not part of the project data)
    # '*.db', '*.sqlite', '*.sqlite3', # Закомментировано, так как .sqlite может быть частью проекта
    # Project specific output
    'collected_files_for_analysis/', # Игнорируем папку, которую мы сами создаем
}

def matches_any_pattern(path: Path, patterns: set) -> bool:
    """Проверяет, соответствует ли путь любому из glob-паттернов."""
    path_str = str(path).replace('\\', '/')
    path_name = path.name
    for pattern in patterns:
        pattern = pattern.rstrip('/')
        if '*' in pattern or '?' in pattern:
            # Это glob паттерн
            try:
                # Проверяем совпадение с именем файла/папки
                if path_str.endswith(path_name) and Path().glob(pattern):
                    # Очень упрощенная проверка. Для надежности лучше использовать fnmatch или pathspec.
                    # Пока просто проверим, если паттерн заканчивается на имя.
                    if pattern.endswith(path_name) or (pattern.startswith('*') and path_name.endswith(pattern[1:])):
                         return True
                # Проверяем, является ли путь подпапкой паттерна-директории
                if pattern.endswith('/') and path_str.startswith(pattern[:-1]):
                    return True
            except:
                pass # Игнорируем ошибки в паттернах
        else:
            # Это просто строка
            if pattern.endswith('/'):
                # Проверяем директорию
                if path_str.startswith(pattern[:-1]) or path_str == pattern[:-1]:
                    return True
            else:
                # Проверяем точное совпадение имени или окончание пути
                if path_name == pattern or path_str.endswith(pattern):
                    return True
    return False

def is_binary(file_path: Path) -> bool:
    """
    Пытается определить, является ли файл двоичным.
    Args:
        file_path (Path): Путь к файлу.
    Returns:
        bool: True, если файл двоичный, False в противном случае.
    """
    # Сначала проверим по расширению (быстро)
    binary_extensions = {
        '.xlsx', '.xls', '.docx', '.doc', '.pdf', '.png', '.jpg', '.jpeg',
        '.gif', '.bmp', '.ico', '.exe', '.dll', '.so', '.dylib', '.class',
        '.zip', '.tar', '.gz', '.rar', '.7z',
        '.mdb', '.accdb' # Другие БД, кроме SQLite
        # .db, .sqlite, .sqlite3 исключены, так как могут быть частью проекта
    }
    if file_path.suffix.lower() in binary_extensions:
        return True

    # Затем попробуем определить по MIME-типу (медленнее, но точнее)
    mime_type, _ = mimetypes.guess_type(str(file_path))
    if mime_type is not None:
        if not mime_type.startswith('text/') and 'charset' not in mime_type:
             if mime_type.startswith('application/'):
                 # application/octet-stream, application/zip точно двоичные
                 if mime_type in ['application/octet-stream', 'application/zip']:
                     return True

    # Если не определили по расширению и MIME, проверим первые байты файла
    try:
        with file_path.open('rb') as f:
            chunk = f.read(1024) # Читаем первые 1024 байта
            if b'\x00' in chunk:
                # Наличие нулевого байта часто указывает на двоичный файл
                return True
            # Попробуем декодировать как текст.
            chunk.decode('utf-8')
    except (UnicodeDecodeError, FileNotFoundError, PermissionError):
        # Если не смогли прочитать или декодировать, считаем двоичным или несуществующим/недоступным
        return True

    return False

def is_likely_text_based(file_path: Path) -> bool:
    """
    Проверяет, является ли файл потенциально текстовым или скриптом/конфигом.
    Это более мягкая проверка, чем is_binary.
    """
    text_like_extensions = {
        '.py', '.pyw', '.pyi', # Python
        '.yaml', '.yml', # YAML
        '.json', # JSON
        '.toml', # TOML
        '.cfg', '.conf', '.config', # Configs
        '.ini', # INI
        '.md', '.markdown', '.rst', '.txt', # Text/Docs
        '.sh', '.bash', '.zsh', # Scripts
        '.bat', '.cmd', # Windows Scripts
        '.sql', # SQL
        '.js', '.jsx', '.ts', '.tsx', # JS/TS
        '.html', '.htm', '.xml', # Markup
        '.css', '.scss', '.sass', # Styles
        '.csv', # CSV (часто текст)
        '.log', # Logs
        '.env', '.env.example', # Env files
        # Добавь сюда другие, если нужно
    }
    return file_path.suffix.lower() in text_like_extensions

def collect_files():
    """Основная логика сбора файлов."""
    project_root = Path(__file__).parent.parent.absolute()
    output_dir = project_root / "collected_files_for_analysis"
    gitignore_path = project_root / ".gitignore"

    print(f"Корневая директория проекта: {project_root}")

    # 1. Создаём/очищаем выходную директорию
    if output_dir.exists():
        shutil.rmtree(output_dir)
        print(f"Очищена предыдущая директория: {output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Создана директория для сбора: {output_dir}")

    # 2. Читаем .gitignore или используем стандартные паттерны
    ignore_patterns = set(DEFAULT_IGNORE_PATTERNS)
    if gitignore_path.exists():
        try:
            with gitignore_path.open('r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        ignore_patterns.add(line)
            print(f"Загружен .gitignore: {gitignore_path}")
        except Exception as e:
            print(f"Ошибка при чтении .gitignore: {e}. Используются стандартные паттерны.")
    else:
        print(f".gitignore не найден. Используются стандартные паттерны игнорирования.")

    # 3. Собираем файлы
    collected_count = 0
    ignored_by_patterns = 0
    ignored_as_binary = 0
    ignored_as_empty = 0
    errors_count = 0

    for file_path in project_root.rglob('*'):
        # Обрабатываем только файлы
        if not file_path.is_file():
            continue

        # Относительный путь от корня проекта
        relative_path = file_path.relative_to(project_root)

        # --- Игнорирование по паттернам ---
        if matches_any_pattern(relative_path, ignore_patterns):
            ignored_by_patterns += 1
            continue

        # --- Игнорирование двоичных файлов ---
        # (Опционально, можно убрать, если хочешь собирать вообще все текстовые)
        if not is_likely_text_based(file_path) and is_binary(file_path):
            ignored_as_binary += 1
            continue

        try:
            # Проверяем, пустой ли файл
            if file_path.stat().st_size == 0:
                ignored_as_empty += 1
                continue

            # Читаем содержимое файла
            try:
                with file_path.open('r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                # Если UTF-8 не сработал, пробуем другие кодировки или пропускаем
                # Для простоты пропустим. Можно улучшить.
                print(f"Не удалось декодировать файл как UTF-8 (пропущен): {relative_path}")
                ignored_as_binary += 1 # Считаем как "не текстовый"
                continue

            if not content.strip(): # Проверяем, не пустая ли строка после strip
                 ignored_as_empty += 1
                 continue

            # Создаём имя файла для копии (заменяем / и \ на _)
            safe_filename = str(relative_path).replace('/', '___').replace('\\', '___')
            output_file_path = output_dir / safe_filename

            # Записываем содержимое в новый файл
            with output_file_path.open('w', encoding='utf-8') as f:
                # Добавляем комментарий в начало файла с его оригинальным путём
                f.write(f"# Оригинальный путь: {relative_path}\n")
                f.write(f"# Размер файла: {file_path.stat().st_size} байт\n")
                f.write("---\n") # Разделитель
                f.write(content)
                f.write("\n") # Добавим новую строку в конце для порядка

            collected_count += 1
            print(f"Собран: {relative_path}")

        except Exception as e:
            print(f"Ошибка при обработке файла {relative_path}: {e}")
            errors_count += 1

    print("\n--- Сводка ---")
    print(f"Собрано файлов: {collected_count}")
    print(f"Проигнорировано (паттерны): {ignored_by_patterns}")
    print(f"Проигнорировано (двоичные): {ignored_as_binary}")
    print(f"Проигнорировано (пустые): {ignored_as_empty}")
    if errors_count > 0:
        print(f"Ошибок: {errors_count}")
    print(f"Файлы собраны в: {output_dir}")

    # 4. Создаём файл с полной структурой проекта
    structure_file = output_dir / "___project_structure.txt"
    try:
        def build_tree(path_obj: Path, prefix: str = "", is_last: bool = True, ignore_set: set = None) -> str:
            """Рекурсивно строит дерево каталогов в виде строки."""
            if not path_obj.exists() or (ignore_set and matches_any_pattern(path_obj.relative_to(project_root), ignore_set)):
                return ""

            display_name = path_obj.name if path_obj.name else str(path_obj)
            line = f"{prefix}{'└── ' if is_last else '├── '}{display_name}{'/' if path_obj.is_dir() else ''}\n"

            if path_obj.is_dir():
                try:
                    children = sorted([p for p in path_obj.iterdir() if not (ignore_set and matches_any_pattern(p.relative_to(project_root), ignore_set))])
                    for i, child in enumerate(children):
                        extension = "    " if is_last else "│   "
                        line += build_tree(child, prefix + extension, i == len(children) - 1, ignore_set)
                except PermissionError:
                     line += f"{prefix}{'└── ' if is_last else '├── '} [Нет доступа]\n"
            return line

        tree_str = f"Структура проекта: {project_root}\n"
        tree_str += f".gitignore: {'Найден и использован' if gitignore_path.exists() else 'Не найден, использованы стандартные паттерны'}\n"
        tree_str += ".\n" # Корень
        try:
            root_children = sorted([p for p in project_root.iterdir() if not (ignore_patterns and matches_any_pattern(p.relative_to(project_root), ignore_patterns))])
            for i, child in enumerate(root_children):
                tree_str += build_tree(child, "", i == len(root_children) - 1, ignore_patterns)
        except PermissionError:
             tree_str += "[Нет доступа к корневой директории]\n"

        with structure_file.open('w', encoding='utf-8') as f:
            f.write(tree_str)
        print(f"Структура проекта сохранена в: {structure_file}")
    except Exception as e:
        print(f"Не удалось создать файл структуры проекта: {e}")


if __name__ == "__main__":
    collect_files()
    print("\nСкрипт завершён.")
