"""
Модуль для определения Pydantic-моделей, которые описывают структуру данных,
передаваемых из Python в Go-экспортер.

Эти модели служат "контрактом" между двумя частями системы и обеспечивают
валидацию данных перед сериализацией в JSON.
"""

from typing import List, Optional, Dict, Any, Union
from pydantic import BaseModel, Field

class ProjectMetadata(BaseModel):
    """Метаданные всего проекта."""
    project_name: str
    author: str
    created_at: str  # ISO 8601

class Formula(BaseModel):
    """Описание формулы в ячейке."""
    cell: str  # Например, "A1"
    formula: str  # Например, "=SUM(B1:C1)"

class Style(BaseModel):
    """
    Описание стиля для диапазона ячеек.
    Структура `style` должна точно соответствовать тому,
    что понимает Go-экспортер (и в конечном итоге Excelize).
    """
    range: str  # Например, "A1:C10"
    style: Dict[str, Any]

class ChartSeries(BaseModel):
    """Описание одной серии данных для диаграммы."""
    name: Optional[str] = None  # Ссылка на ячейку с названием, например "Sheet1!$A$1"
    categories: Optional[str] = None  # Диапазон категорий, например "Sheet1!$B$1:$D$1"
    values: str  # Диапазон значений, например "Sheet1!$B$2:$D$2"

class Chart(BaseModel):
    """Описание диаграммы."""
    type: str  # Тип диаграммы, понятный Excelize (например, "col", "line", "pie")
    position: str  # Ячейка, где будет размещена диаграмма, например "E5"
    title: Optional[str] = None
    series: List[ChartSeries]

class SheetData(BaseModel):
    """Данные для одного листа Excel."""
    name: str
    # Данные представлены как список строк, где каждая строка - список ячеек.
    # Пустая ячейка представляется как None.
    data: List[List[Optional[str]]]
    formulas: List[Formula] = Field(default_factory=list)
    styles: List[Style] = Field(default_factory=list)
    charts: List[Chart] = Field(default_factory=list)

class ExportData(BaseModel):
    """Корневая модель для всего экспортируемого проекта."""
    metadata: ProjectMetadata
    sheets: List[SheetData]