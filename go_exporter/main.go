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
// Внимание: Поддерживаются только базовые типы, доступные в Excelize v2.9.1.
// Сложные типы (например, ColStacked, LinePercentStacked, Pie3D, DoughnutExploded)
// могут требовать обновления Excelize или более сложной реализации.
func convertChartType(chartTypeStr string) excelize.ChartType {
	switch chartTypeStr {
	case "col":
		return excelize.Col
	case "col3D": // Доступен в v2.9.1
		return excelize.Col3D
	case "line":
		return excelize.Line
	case "pie":
		return excelize.Pie
	case "pie3D": // Доступен в v2.9.1
		return excelize.Pie3D
	case "doughnut":
		return excelize.Doughnut
	case "bar": // Добавлен для полноты, если используется
		return excelize.Bar
	case "area": // Добавлен для полноты, если используется
		return excelize.Area
	case "radar": // Добавлен для полноты, если используется
		return excelize.Radar
	case "scatter": // Добавлен для полноты, если используется
		return excelize.Scatter
	case "stock": // Добавлен для полноты, если используется
		return excelize.Stock
	case "combo": // Добавлен для полноты, если используется
		return excelize.Combo
	// Следующие типы НЕ поддерживаются в Excelize v2.9.1 и возвращают базовый тип 'Col'
	case "colStacked", "colPercentStacked", "col3DClustered", "col3DStacked", "col3DPercentStacked",
		"lineStacked", "linePercentStacked", "line3D", "pieOfPie", "barOfPie", "doughnutExploded":
		fmt.Printf("Warning: Chart type '%s' is not supported in Excelize v2.9.1, using 'col' as default.\n", chartTypeStr)
		return excelize.Col
	default:
		// Возвращаем тип по умолчанию, если тип не распознан
		fmt.Printf("Warning: Unknown chart type '%s', using 'col' as default.\n", chartTypeStr)
		return excelize.Col // или другой тип по умолчанию
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
	for _, sheet := range exportData.Sheets {
		// Создание нового листа
		if err := f.SetSheetName(f.GetSheetName(0), sheet.Name); err != nil {
			log.Printf("Warning: could not rename first sheet to '%s': %v", sheet.Name, err)
		}

		// Заполнение данными
		for rowIndex, row := range sheet.Data {
			cellRow := rowIndex + 1
			for colIndex, cellValue := range row {
				cellCol := colIndex + 1
				cellName, _ := excelize.ColumnNumberToName(cellCol)
				cellName += fmt.Sprintf("%d", cellRow)

				if cellValue != nil {
					f.SetCellValue(sheet.Name, cellName, *cellValue)
				}
			}
		}

		// Добавление формул
		for _, formula := range sheet.Formulas {
			f.SetCellFormula(sheet.Name, formula.Cell, formula.Formula)
		}

		// TODO: Реализовать применение стилей
		// Это будет самая сложная часть, так как нужно отобразить
		// структуру стилей из Python в формат Excelize.

		// Добавление диаграмм
		for _, chart := range sheet.Charts {
			chartConfig := &excelize.Chart{
				Type: convertChartType(chart.Type), // <-- Изменённая строка
				Series: []excelize.ChartSeries{},
				Title:  []excelize.RichTextRun{{Text: chart.Title}},
			}

			for _, series := range chart.Series {
				chartConfig.Series = append(chartConfig.Series, excelize.ChartSeries{
					Name:       series.Name,
					Categories: series.Categories,
					Values:     series.Values,
				})
			}

			if err := f.AddChart(sheet.Name, chart.Position, chartConfig); err != nil {
				log.Printf("Warning: could not add chart at %s: %v", chart.Position, err)
			}
		}

		// Если есть ещё листы, создадим их
		// (Первый лист уже существует по умолчанию)
		// ... (логика для создания дополнительных листов)
	}

	// Сохранение файла
	if err := f.SaveAs(*outputFile); err != nil {
		log.Fatalf("Error saving file: %v", err)
	}

	fmt.Printf("Successfully exported to %s\n", *outputFile)
}