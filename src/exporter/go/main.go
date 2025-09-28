package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

// ExportData структура для всего экспортируемого проекта
type ExportData struct {
	Metadata ProjectMetadata `json:"metadata"`
	Sheets   []SheetData     `json:"sheets"`
}

// ProjectMetadata метаданные проекта
type ProjectMetadata struct {
	ProjectName string `json:"project_name"`
	Author      string `json:"author"`
	CreatedAt   string `json:"created_at"`
}

// SheetData структура для данных одного листа
type SheetData struct {
	Name    string       `json:"name"`
	Data    [][]*string  `json:"data"` // nil для пустых ячеек
	Formulas []Formula   `json:"formulas,omitempty"`
	Styles   []Style     `json:"styles,omitempty"`
	Charts   []Chart     `json:"charts,omitempty"`
	MergedCells []string `json:"merged_cells,omitempty"`
}

// Formula структура для формулы
type Formula struct {
	Cell    string `json:"cell"`
	Formula string `json:"formula"`
}

// Style структура для стиля (упрощённая для примера)
// В реальной реализации нужно будет отобразить все свойства стилей из Python
type Style struct {
	Range string                 `json:"range"`
	Style map[string]interface{} `json:"style"`
}

// Chart структура для диаграммы
type Chart struct {
	Type     string       `json:"type"`
	Position string       `json:"position"`
	Title    string       `json:"title,omitempty"`
	Series   []ChartSeries `json:"series"`
}

// ChartSeries структура для серии диаграммы
type ChartSeries struct {
	Name       string `json:"name"`
	Categories string `json:"categories"`
	Values     string `json:"values"`
}

// convertChartType преобразует строку типа диаграммы из JSON в excelize.ChartType.
// Поддерживаются только базовые типы, доступные в Excelize v2.9.1.
// Неподдерживаемые или неизвестные типы возвращают базовый тип 'Col'.
func convertChartType(chartTypeStr string) excelize.ChartType {
	switch chartTypeStr {
	// Поддерживаемые типы в v2.9.1
	case "col":
		return excelize.Col
	case "line":
		return excelize.Line
	case "pie":
		return excelize.Pie
	case "bar":
		return excelize.Bar
	case "area":
		return excelize.Area
	case "scatter":
		return excelize.Scatter
	case "doughnut":
		return excelize.Doughnut
	// Типы, которые могут быть в JSON, но не поддерживаются напрямую в v2.9.1.
	// Возвращаем наиболее подходящий базовый тип или 'Col' по умолчанию.
	// Это предотвращает ошибки компиляции.
	case "colStacked", "colPercentStacked", "col3D", "col3DClustered", "col3DStacked", "col3DPercentStacked",
		"lineStacked", "linePercentStacked", "line3D", "pie3D", "pieOfPie", "barOfPie", "doughnutExploded":
		// Можно логировать предупреждение, если нужно
		// fmt.Printf("Warning: Chart type '%s' is not directly supported in Excelize v2.9.1, using 'col' as fallback.\n", chartTypeStr)
		return excelize.Col
	default:
		// Неизвестный тип - возвращаем 'Col' как тип по умолчанию
		// Лучше логировать это как предупреждение
		fmt.Printf("Warning: Unknown chart type '%s', using 'col' as default.\n", chartTypeStr)
		return excelize.Col
	}
}

func main() {
	// Парсинг аргументов командной строки
	inputFile := flag.String("input", "", "Path to the input JSON file")
	outputFile := flag.String("output", "", "Path to the output XLSX file")
	flag.Parse()

	if *inputFile == "" || *outputFile == "" {
		fmt.Println("Usage: go_excel_exporter -input <input.json> -output <output.xlsx>")
		os.Exit(1)
	}

	// Чтение JSON-файла
	jsonData, err := os.ReadFile(*inputFile)
	if err != nil {
		log.Fatalf("Error reading input file: %v", err)
	}

	// Парсинг JSON в структуру Go
	var exportData ExportData
	err = json.Unmarshal(jsonData, &exportData)
	if err != nil {
		log.Fatalf("Error parsing JSON: %v", err)
	}

	// Создание нового Excel-файла
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("Error closing file: %v", err)
		}
	}()

	// Обработка каждого листа
	for i, sheet := range exportData.Sheets {
		var sheetName string
		if i == 0 {
			// Переименовываем первый (дефолтный) лист
			sheetName = sheet.Name
			if err := f.SetSheetName("Sheet1", sheetName); err != nil {
				log.Printf("Warning: could not rename default sheet to '%s': %v", sheetName, err)
				// Если не удалось переименовать, используем имя по умолчанию или генерируем новое?
				// Для простоты, продолжим с sheet.Name, надеясь, что SetSheetName создаст его, если он не первый.
				// Более надежный способ - проверить существование и создать/переименовать.
			}
		} else {
			// Для последующих листов создаем новый
			_, err := f.NewSheet(sheet.Name) // <-- Исправлено: игнорируем индекс листа
			if err != nil {
				log.Printf("Warning: could not create new sheet '%s': %v", sheet.Name, err)
				continue // Пропускаем этот лист, если не удалось создать
			}
			sheetName = sheet.Name
			// Убедимся, что активный лист остается первым или последним созданным?
			// f.SetActiveSheet(index) // Опционально
		}

		// Заполнение данными
		for rowIndex, row := range sheet.Data {
			for colIndex, cellValue := range row {
				// Excelize использует 1-индексацию
				cellRow := rowIndex + 1
				cellCol := colIndex + 1
				// Преобразуем номер столбца в имя (A, B, ..., Z, AA, AB, ...)
				cellName, err := excelize.ColumnNumberToName(cellCol)
				if err != nil {
					log.Printf("Error converting column number %d to name: %v", cellCol, err)
					continue
				}
				cellAddress := fmt.Sprintf("%s%d", cellName, cellRow)

				if cellValue != nil {
					// Устанавливаем значение ячейки
					// f.SetCellValue(sheetName, cellAddress, *cellValue) // Этот способ тоже работает
					// Используем более конкретный метод, если известен тип, но для общего случая SetCellValue подходит.
					if err := f.SetCellValue(sheetName, cellAddress, *cellValue); err != nil {
						log.Printf("Warning: could not set cell value at %s on sheet '%s': %v", cellAddress, sheetName, err)
					}
				}
			}
		}

		// Добавление формул
		for _, formula := range sheet.Formulas {
			if err := f.SetCellFormula(sheetName, formula.Cell, formula.Formula); err != nil {
				log.Printf("Warning: could not set formula at %s on sheet '%s': %v", formula.Cell, sheetName, err)
			}
		}

		// TODO: Реализовать применение стилей
		// Это будет самая сложная часть, так как нужно отобразить
		// структуру стилей из Python в формат Excelize.

		// Добавление диаграмм
		for _, chart := range sheet.Charts {
			// Создаем конфигурацию диаграммы
			chartConfig := &excelize.Chart{
				Type: convertChartType(chart.Type), // <-- Используем нашу функцию
				// Series будет заполнен ниже
				Series: []excelize.ChartSeries{},
				// Title теперь принимает []excelize.RichTextRun
				Title: []excelize.RichTextRun{{Text: chart.Title}},
			}

			// Заполняем серии данных для диаграммы
			for _, series := range chart.Series {
				chartConfig.Series = append(chartConfig.Series, excelize.ChartSeries{
					Name:       series.Name,
					Categories: series.Categories,
					Values:     series.Values,
				})
			}

			// Добавляем диаграмму на лист
			if err := f.AddChart(sheetName, chart.Position, chartConfig); err != nil {
				log.Printf("Warning: could not add chart at %s on sheet '%s': %v", chart.Position, sheetName, err)
			}
		}
		// Конец обработки диаграмм для текущего листа

		// Применение объединенных ячеек
		for _, mergedCellRange := range sheet.MergedCells {
			if err := f.MergeCell(sheetName, mergedCellRange); err != nil {
				log.Printf("Warning: could not merge cells '%s' on sheet '%s': %v", mergedCellRange, sheetName, err)
			}
		}
		// Конец применения объединенных ячеек для текущего листа

		// TODO: Обработка дополнительных элементов (изображения, таблицы и т.д.)
	}
	// Конец обработки всех листов

	// Сохранение файла
	if err := f.SaveAs(*outputFile); err != nil {
		log.Fatalf("Error saving file: %v", err)
	}

	fmt.Printf("Successfully exported to %s\n", *outputFile)
}