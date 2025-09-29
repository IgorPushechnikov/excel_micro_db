# backend/api/fastapi_server.py
"""
Модуль для запуска HTTP-сервера на основе FastAPI.
Обеспечивает API для взаимодействия с GUI (Tauri).
"""

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
import uvicorn
import logging
import sys
from pathlib import Path

# Добавляем backend в путь, чтобы можно было импортировать core, storage и т.д.
# Предполагается, что этот файл будет запускаться из backend/ или что PYTHONPATH настроен соответствующим образом.
# Или main.py добавит backend в sys.path перед импортом этого модуля.
# sys.path.insert(0, str(Path(__file__).parent.parent)) # Можно раскомментировать при прямом запуске

from core.app_controller import create_app_controller
from utils.logger import get_logger

# Получаем логгер для этого модуля
logger = get_logger(__name__)

# --- Pydantic модели для запросов/ответов (примеры) ---

class AnalyzeRequest(BaseModel):
    """Модель для запроса на анализ Excel-файла."""
    excel_file_path: str
    project_path: str
    options: Optional[Dict[str, Any]] = None # Дополнительные опции анализа


class AnalyzeResponse(BaseModel):
    """Модель для ответа на запрос анализа."""
    success: bool
    message: str
    sheets: Optional[List[str]] = None # Список имён листов, например


class ExportRequest(BaseModel):
    """Модель для запроса на экспорт проекта."""
    export_type: str # 'excel', 'go_excel' и т.д.
    output_path: str
    project_path: str


class ExportResponse(BaseModel):
    """Модель для ответа на запрос экспорта."""
    success: bool
    message: str
    exported_file_path: Optional[str] = None


class SheetsResponse(BaseModel):
    """Модель для ответа на запрос списка листов."""
    sheets: List[str]


# --- Создание экземпляра FastAPI ---

app = FastAPI(
    title="Excel Micro DB API",
    description="API для взаимодействия с Excel Micro DB.",
    version="0.1.0",
)


# --- Эндпоинты ---

@app.post("/api/analyze", response_model=AnalyzeResponse)
async def api_analyze(request: AnalyzeRequest):
    """Анализ Excel-файла и сохранение результатов в проект."""
    logger.info(f"Получен запрос на анализ: {request.excel_file_path} в проект {request.project_path}")
    try:
        # --- Создание и инициализация AppController ---
        # TODO: Управление жизненным циклом AppController (один на сервер? на запрос?)
        # Для начала, создадим новый для каждого запроса.
        app_controller = create_app_controller(project_path=request.project_path)
        if not app_controller.initialize():
             logger.error("Не удалось инициализировать AppController для анализа.")
             raise HTTPException(status_code=500, detail="Ошибка инициализации приложения")

        # Проверяем, загружен ли проект
        if not app_controller.is_project_loaded:
             logger.error("Проект не загружен для анализа.")
             raise HTTPException(status_code=400, detail="Проект не загружен")

        # --- Вызов логики анализа через AppController ---
        # TODO: Реализовать вызов app_controller.analyze_excel_file с переданными параметрами
        logger.debug(f"Вызов AppController для анализа с опциями: {request.options}")
        success = app_controller.analyze_excel_file(request.excel_file_path, options=request.options or {})

        if success:
            # Попробуем получить список листов после анализа
            # TODO: Реализовать метод в AppController для получения списка листов проекта
            # sheets_list = app_controller.get_project_sheet_names() # Предполагаемый метод
            sheets_list = [] # Заглушка
            logger.info("Анализ успешно завершён.")
            return AnalyzeResponse(success=True, message="Анализ завершён", sheets=sheets_list)
        else:
            logger.error("Ошибка при анализе файла через AppController.")
            raise HTTPException(status_code=500, detail="Ошибка анализа файла")

    except HTTPException:
        # Переподнимаем HTTPException, чтобы FastAPI корректно её обработал
        raise
    except Exception as e:
        logger.error(f"Неожиданная ошибка при анализе: {e}")
        # Логируем traceback для отладки
        import traceback
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Внутренняя ошибка сервера: {e}")


@app.post("/api/export", response_model=ExportResponse)
async def api_export(request: ExportRequest):
    """Экспорт проекта в указанный формат."""
    logger.info(f"Получен запрос на экспорт: тип '{request.export_type}', в {request.output_path}, из проекта {request.project_path}")
    try:
        # --- Создание и инициализация AppController ---
        app_controller = create_app_controller(project_path=request.project_path)
        if not app_controller.initialize():
             logger.error("Не удалось инициализировать AppController для экспорта.")
             raise HTTPException(status_code=500, detail="Ошибка инициализации приложения")

        # Проверяем, загружен ли проект
        if not app_controller.is_project_loaded:
             logger.error("Проект не загружен для экспорта.")
             raise HTTPException(status_code=400, detail="Проект не загружен")

        # --- Вызов логики экспорта через AppController ---
        # TODO: Реализовать вызов app_controller.export_results с переданными параметрами
        logger.debug(f"Вызов AppController для экспорта типа {request.export_type}")
        success = app_controller.export_results(export_type=request.export_type, output_path=request.output_path)

        if success:
            logger.info("Экспорт успешно завершён.")
            return ExportResponse(success=True, message="Экспорт завершён", exported_file_path=request.output_path)
        else:
            logger.error("Ошибка при экспорте через AppController.")
            raise HTTPException(status_code=500, detail="Ошибка экспорта")

    except HTTPException:
        # Переподнимаем HTTPException, чтобы FastAPI корректно её обработал
        raise
    except Exception as e:
        logger.error(f"Неожиданная ошибка при экспорте: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Внутренняя ошибка сервера: {e}")


@app.get("/api/sheets", response_model=SheetsResponse)
async def api_get_sheets(project_path: str):
    """Получение списка листов из загруженного проекта."""
    logger.info(f"Получен запрос на получение листов для проекта: {project_path}")
    try:
        # --- Создание и инициализация AppController ---
        app_controller = create_app_controller(project_path=project_path)
        if not app_controller.initialize():
             logger.error("Не удалось инициализировать AppController для получения листов.")
             raise HTTPException(status_code=500, detail="Ошибка инициализации приложения")

        # Проверяем, загружен ли проект
        if not app_controller.is_project_loaded:
             logger.error("Проект не загружен для получения листов.")
             raise HTTPException(status_code=400, detail="Проект не загружен")

        # --- Вызов логики получения листов через AppController ---
        # TODO: Реализовать метод в AppController для получения списка листов проекта
        # sheets_list = app_controller.get_project_sheet_names()
        sheets_list = [] # Заглушка
        logger.info(f"Получен список листов (заглушка): {sheets_list}")
        return SheetsResponse(sheets=sheets_list)

    except HTTPException:
        # Переподнимаем HTTPException, чтобы FastAPI корректно её обработал
        raise
    except Exception as e:
        logger.error(f"Неожиданная ошибка при получении листов: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Внутренняя ошибка сервера: {e}")


# --- Функция для запуска сервера ---

def run_server(host: str = "127.0.0.1", port: int = 8000):
    """
    Запускает FastAPI-сервер с помощью uvicorn.

    Args:
        host (str): Хост для запуска сервера.
        port (int): Порт для запуска сервера.
    """
    logger.info(f"Запуск FastAPI-сервера на {host}:{port}")
    # Используем строку модуля и имя приложения для uvicorn
    # Предполагаем, что файл запускается как python -m backend.api.fastapi_server или uvicorn запускает его
    # Если запускать через python backend/api/fastapi_server.py, то нужно указать app напрямую
    uvicorn.run(app, host=host, port=port, log_level="info", reload=False) # reload=True только для разработки

# Альтернативная точка входа, если файл запускается напрямую
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Запуск FastAPI-сервера Excel Micro DB.")
    parser.add_argument("--host", default="127.0.0.1", help="Хост для запуска сервера")
    parser.add_argument("--port", type=int, default=8000, help="Порт для запуска сервера")
    args = parser.parse_args()

    run_server(args.host, args.port)
