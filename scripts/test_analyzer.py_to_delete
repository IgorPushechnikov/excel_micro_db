# scripts/test_analyzer.py
"""
Скрипт для тестирования анализатора логики Excel.
"""

import sys
from pathlib import Path

# Добавляем корень проекта в путь поиска модулей
# project_root = Path(__file__).parent.absolute() # Старая строка
project_root = Path(__file__).parent.parent.absolute() # Поднимаемся на уровень выше из scripts/ в корень проекта
sys.path.insert(0, str(project_root))

from src.analyzer.logic_documentation import create_documentation_from_excel
from src.utils.logger import get_logger

# Получаем логгер для этого скрипта
logger = get_logger(__name__)

def main():
    """Основная функция тестирования анализатора."""
    logger.info("Начало тестирования анализатора логики Excel")
    
    # Пути к тестовым файлам
    test_excel_file = project_root / "data" / "samples" / "test_sample.xlsx"
    test_project_path = project_root / "test_workspace" / "analyzer_test_project"
    
    # Проверяем существование тестового Excel файла
    if not test_excel_file.exists():
        logger.error(f"Тестовый Excel файл не найден: {test_excel_file}")
        logger.info("Пожалуйста, сначала создайте тестовый файл с помощью scripts/create_test_excel.py")
        return False
    
    try:
        # Создаем тестовый проект для сохранения результатов
        test_project_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"Тестовый проект создан в: {test_project_path}")
        
        # Опции для анализа
        options = {
            'max_rows': 1000,
            'include_formulas': True
        }
        
        # Запускаем анализатор
        logger.info(f"Анализ файла: {test_excel_file}")
        documentation = create_documentation_from_excel(
            excel_file_path=str(test_excel_file),
            project_path=str(test_project_path),
            options=options
        )
        
        # Выводим краткую информацию о результатах
        if documentation:
            logger.info("Анализ завершен успешно!")
            logger.info(f"Версия документации: {documentation.get('version', 'Не указана')}")
            logger.info(f"Источник: {documentation.get('source_file', 'Не указан')}")
            logger.info(f"Количество листов: {len(documentation.get('sheets', []))}")
            
            # Выводим информацию о каждом листе
            for sheet in documentation.get('sheets', []):
                sheet_name = sheet.get('name', 'Без названия')
                rows_count = sheet.get('rows_count', 0)
                cols_count = sheet.get('cols_count', 0)
                formulas_count = len(sheet.get('formulas', []))
                cross_sheet_refs_count = len(sheet.get('cross_sheet_references', []))
                logger.info(f"  Лист '{sheet_name}': {rows_count} строк, {cols_count} столбцов, {formulas_count} формул, {cross_sheet_refs_count} межл. ссылок")
                
                # Выводим информацию о формулах на листе (первые 2)
                formulas_on_sheet = sheet.get('formulas', [])
                if formulas_on_sheet:
                    logger.info(f"    Формулы на листе '{sheet_name}' (первые 2):")
                    for formula in formulas_on_sheet[:2]:
                        cell_addr = formula.get('cell_address', 'Неизвестно')
                        formula_text = formula.get('formula_text', 'Неизвестно')
                        result_type = formula.get('result_type', 'unknown')
                        logger.info(f"      {cell_addr}: {formula_text} [Тип: {result_type}]")
                
                # Выводим информацию о межлистовых ссылках на листе (первые 2)
                cross_sheet_refs_on_sheet = sheet.get('cross_sheet_references', [])
                if cross_sheet_refs_on_sheet:
                    logger.info(f"    Межлистовые ссылки на листе '{sheet_name}' (первые 2):")
                    for ref in cross_sheet_refs_on_sheet[:2]:
                        formula_cell = ref.get('formula_cell', 'Неизвестно')
                        target_sheet = ref.get('target_sheet', 'Неизвестно')
                        target_cell = ref.get('target_cell', 'Неизвестно')
                        logger.info(f"      {formula_cell} -> '{target_sheet}'!{target_cell}")
            
            # Выводим сводную информацию
            summary = documentation.get('summary', {})
            logger.info(f"Сводка:")
            logger.info(f"  Всего строк данных: {summary.get('total_data_rows', 0)}")
            logger.info(f"  Всего формул: {summary.get('total_formulas', 0)}")
            logger.info(f"  Наиболее частый тип данных: {summary.get('most_common_data_type', 'Не определен')}")
            logger.info(f"  Оценка качества данных: {summary.get('data_quality_score', 0)}/100")
            
            logger.info(f"Результаты сохранены в проекте: {test_project_path}")
            return True
        else:
            logger.error("Анализ не вернул результатов")
            return False
            
    except Exception as e:
        logger.error(f"Ошибка при тестировании анализатора: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return False

if __name__ == "__main__":
    success = main()
    if success:
        logger.info("Тестирование анализатора завершено успешно")
        sys.exit(0)
    else:
        logger.error("Тестирование анализатора завершено с ошибками")
        sys.exit(1)