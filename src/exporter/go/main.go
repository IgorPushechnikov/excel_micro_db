// SPDX-License-Identifier: MIT
// Copyright (c) 2025 Ivan Pushkin
// All rights reserved.

// Package main provides the Go-based Excel exporter for the Excel Micro DB project.
// It reads a JSON file containing project data (sheets, cells, formulas, styles, charts, merged cells)
// and generates an .xlsx file using the excelize library.
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

// ExportData represents the root structure of the project data to be exported.
type ExportData struct {
	Metadata ProjectMetadata `json:"metadata"`
	Sheets   []SheetData     `json:"sheets"`
}

// ProjectMetadata holds project-level metadata.
type ProjectMetadata struct {
	ProjectName string `json:"project_name"`
	Author      string `json:"author"`
	CreatedAt   string `json:"created_at"`
}

// SheetData holds the data for a single worksheet.
type SheetData struct {
	Name         string       `json:"name"`
	Data         [][]*string  `json:"data"` // nil represents empty cells
	Formulas     []Formula    `json:"formulas,omitempty"`
	Styles       []Style      `json:"styles,omitempty"`
	Charts       []Chart      `json:"charts,omitempty"`
	MergedCells  []string     `json:"merged_cells,omitempty"`
}

// Formula represents a cell formula.
type Formula struct {
	Cell    string `json:"cell"`
	Formula string `json:"formula"`
}

// Style represents a cell style definition.
type Style struct {
	Range string                 `json:"range"` // e.g., "A1:B10"
	Style map[string]interface{} `json:"style"` // Style attributes dictionary
}

// Chart represents a chart definition.
type Chart struct {
	Type     string        `json:"type"`
	Position string        `json:"position"`
	Title    string        `json:"title,omitempty"`
	Series   []ChartSeries `json:"series"`
}

// ChartSeries represents a data series for a chart.
type ChartSeries struct {
	Name       string `json:"name"`
	Categories string `json:"categories"`
	Values     string `json:"values"`
}
// convertChartType converts a string chart type from JSON to excelize.ChartType.
// It supports basic types available in Excelize v2.9.1.
// Unsupported or unknown types default to 'Col'.
func convertChartType(chartTypeStr string) excelize.ChartType {
	switch chartTypeStr {
	// Supported types in v2.9.1
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
	// Types that might be in JSON but are not directly supported in v2.9.1.
	// Return the most suitable basic type or 'Col' as a fallback.
	// This prevents compilation errors.
	case "colStacked", "colPercentStacked", "col3D", "col3DClustered", "col3DStacked", "col3DPercentStacked",
		"lineStacked", "linePercentStacked", "line3D", "pie3D", "pieOfPie", "barOfPie", "doughnutExploded":
		// A warning could be logged if needed
		// fmt.Printf("Warning: Chart type '%s' is not directly supported in Excelize v2.9.1, using 'col' as fallback.\n", chartTypeStr)
		return excelize.Col
	default:
		// Unknown type - default to 'Col'
		// It's better to log this as a warning
		fmt.Printf("Warning: Unknown chart type '%s', using 'col' as default.\n", chartTypeStr)
		return excelize.Col
	}
}

// convertStyleToExcelizeOptions converts a map[string]interface{} style definition from JSON into an *excelize.Style.
// This is the core "translator" function for styles to excelize.
// Returns *excelize.Style and an error if conversion fails.
func convertStyleToExcelizeOptions(styleMap map[string]interface{}) (*excelize.Style, error) {
	// Create an empty excelize style structure
	excelizeStyle := &excelize.Style{}

	// Process nested structures: font, fill, alignment, border
	if fontData, ok := styleMap["font"].(map[string]interface{}); ok {
		// Excelize.Font fields: Bold, Italic, Underline, Family, Size, Color, etc.
		// Name -> Family
		if name, ok := fontData["name"].(string); ok {
			excelizeStyle.Font = &excelize.Font{
				Family: name,
			}
		} else {
			excelizeStyle.Font = &excelize.Font{} // Initialize if not present
		}

		if b, ok := fontData["b"].(bool); ok {
			excelizeStyle.Font.Bold = b
		}
		if i, ok := fontData["i"].(bool); ok {
			excelizeStyle.Font.Italic = i
		}
		if colorData, ok := fontData["color"].(map[string]interface{}); ok {
			if rgb, ok := colorData["rgb"].(string); ok && len(rgb) == 6 {
				excelizeStyle.Font.Color = rgb
			}
		}
		if sz, ok := fontData["sz"].(float64); ok { // JSON numbers are float64
			excelizeStyle.Font.Size = sz
		}
	}

	if fillData, ok := styleMap["fill"].(map[string]interface{}); ok {
		// Excelize.Fill fields: Type, Pattern, Color, Shading
		// BgColor/FgColor -> Color
		// patternType -> Pattern
		if patternType, ok := fillData["patternType"].(string); ok {
			// Map patternType to excelize.Pattern (int)
			var fillPattern int
			switch patternType {
			case "solid":
				fillPattern = 1 // solid
			case "darkGray":
				fillPattern = 2 // darkGray
			case "mediumGray":
				fillPattern = 3 // mediumGray
			case "lightGray":
				fillPattern = 4 // lightGray
			case "gray125":
				fillPattern = 17 // gray125
			case "gray0625":
				fillPattern = 18 // gray0625
			default:
				fillPattern = 0 // none
			}
			excelizeStyle.Fill = excelize.Fill{
				Type:    "pattern",
				Pattern: fillPattern,
			}
		} else {
			excelizeStyle.Fill.Type = "pattern" // Default to pattern
		}

		// Handle bgColor and fgColor - use the first one found or combine if needed
		// For simplicity, let's take bgColor first, then fgColor if bgColor is not set
		// In excelize, Fill.Color for pattern fills is a []string
		var fillColor string
		if bgColorData, ok := fillData["bgColor"].(map[string]interface{}); ok {
			if rgb, ok := bgColorData["rgb"].(string); ok && len(rgb) == 6 {
				fillColor = rgb
			}
		}
		if fillColor == "" {
			if fgColorData, ok := fillData["fgColor"].(map[string]interface{}); ok {
				if rgb, ok := fgColorData["rgb"].(string); ok && len(rgb) == 6 {
					fillColor = rgb
				}
			}
		}
		if fillColor != "" {
			excelizeStyle.Fill.Color = []string{fillColor} // For pattern fill, Color is a slice of strings
		}
	}

	if alignmentData, ok := styleMap["alignment"].(map[string]interface{}); ok {
		// Excelize.Alignment fields: Horizontal, Vertical, WrapText, TextRotation, etc.
		excelizeStyle.Alignment = &excelize.Alignment{}
		if horizontal, ok := alignmentData["horizontal"].(string); ok {
			excelizeStyle.Alignment.Horizontal = horizontal // "left", "center", "right", etc.
		}
		if vertical, ok := alignmentData["vertical"].(string); ok {
			excelizeStyle.Alignment.Vertical = vertical // "top", "middle", "bottom", etc.
		}
		// textRotation might be int or float64 in JSON
		if rotation, ok := alignmentData["textRotation"].(float64); ok {
			excelizeStyle.Alignment.TextRotation = int(rotation)
		}
	}

	if borderData, ok := styleMap["border"].(map[string]interface{}); ok {
		// Excelize expects a slice of Border structs
		// Check each side (top, bottom, left, right)
		// Map string values from openpyxl to excelize integer codes
		// See excelize documentation: https://pkg.go.dev/github.com/xuri/excelize/v2#FormatBorder
		if topData, ok := borderData["top"].(map[string]interface{}); ok {
			styleCode := getStyleFromMap(topData)
			color := getColorFromMap(topData)
			excelizeStyle.Border = append(excelizeStyle.Border, excelize.Border{
				Type:  "top",
				Color: color, // Color is a string
				Style: styleCode,
			})
		}
		if bottomData, ok := borderData["bottom"].(map[string]interface{}); ok {
			styleCode := getStyleFromMap(bottomData)
			color := getColorFromMap(bottomData)
			excelizeStyle.Border = append(excelizeStyle.Border, excelize.Border{
				Type:  "bottom",
				Color: color,
				Style: styleCode,
			})
		}
		if leftData, ok := borderData["left"].(map[string]interface{}); ok {
			styleCode := getStyleFromMap(leftData)
			color := getColorFromMap(leftData)
			excelizeStyle.Border = append(excelizeStyle.Border, excelize.Border{
				Type:  "left",
				Color: color,
				Style: styleCode,
			})
		}
		if rightData, ok := borderData["right"].(map[string]interface{}); ok {
			styleCode := getStyleFromMap(rightData)
			color := getColorFromMap(rightData)
			excelizeStyle.Border = append(excelizeStyle.Border, excelize.Border{
				Type:  "right",
				Color: color,
				Style: styleCode,
			})
		}
	}

	// number_format
	if numFmt, ok := styleMap["number_format"].(string); ok {
		// NumFmt in excelize.Style is an int, not string.
		// We need to map string formats to int codes or use a different approach.
		// For now, let's assume a direct mapping might be complex and log a warning.
		// A common approach is to use excelize.SetColStyle/SetRowStyle/SetCellStyle with a predefined style ID.
		// However, for simplicity, we might need to create a mapping table or use SetCellStyle with format strings.
		// Let's try to see if excelize supports setting NumFmt directly from string via NewStyle options.
		// Actually, excelize.NewStyle *does* accept NumFmt as a string key in the options map.
		// But in the Style struct, it's an int.
		// Let's map some common ones or use a generic approach.
		// This is a common challenge when mapping from openpyxl (string) to excelize (int).
		// For now, we'll set it as int if it's a known code, otherwise log.
		// A better approach would be to use a map or handle this in the Python side.
		// For this example, let's try to parse the string or use a default.
		// Let's use the int directly from the string if it's a number, or handle common cases.
		// Actually, excelize.Style struct does have NumFmt as int. This is tricky.
		// Let's assume the Python side sends the integer code, or we map it here.
		// Let's map some common string formats to excelize codes.
		// General = 0, 0 = 1, 0.00 = 2, #,##0 = 3, #,##0.00 = 4, 0% = 9, 0.00% = 10, 0.00E+00 = 11, # ?/? = 12, # ??/?? = 13, mm-dd-yy = 14, d-mmm-yy = 15, d-mmm = 16, mmm-yy = 17, h:mm AM/PM = 18, h:mm:ss AM/PM = 19, h:mm = 20, h:mm:ss = 21, m/d/yy h:mm = 22, [Red] #,##0.00 = 37, [Red] #,##0;[Green] -#,##0 = 38, [Red] #,##0.00;[Green] -#,##0.00 = 39, [Red] #,##0;[Green] -#,##0 = 40, [Red] #,##0.00;[Green] -#,##0.00 = 41, [Red] #,##0;[Green] -#,##0 = 42, [Red] #,##0.00;[Green] -#,##0.00 = 43, [Red] #,##0;[Green] -#,##0 = 44, [Red] #,##0.00;[Green] -#,##0.00 = 45, [Red] #,##0;[Green] -#,##0 = 46, [Red] #,##0.00;[Green] -#,##0.00 = 47, [Red] #,##0;[Green] -#,##0 = 48, [Red] #,##0.00;[Green] -#,##0.00 = 49, @ = 50
		// Let's try to map common string formats to int codes.
		// This is a simplification. A more robust solution would be to pass the int code from Python.
		var numFmtCode int
		switch numFmt {
		case "General":
			numFmtCode = 0
		case "0":
			numFmtCode = 1
		case "0.00":
			numFmtCode = 2
		case "#,##0":
			numFmtCode = 3
		case "#,##0.00":
			numFmtCode = 4
		case "0%":
			numFmtCode = 9
		case "0.00%":
			numFmtCode = 10
		case "0.00E+00":
			numFmtCode = 11
		case "# ?/?":
			numFmtCode = 12
		case "# ??/??":
			numFmtCode = 13
		case "mm-dd-yy":
			numFmtCode = 14
		case "d-mmm-yy":
			numFmtCode = 15
		case "d-mmm":
			numFmtCode = 16
		case "mmm-yy":
			numFmtCode = 17
		case "h:mm AM/PM":
			numFmtCode = 18
		case "h:mm:ss AM/PM":
			numFmtCode = 19
		case "h:mm":
			numFmtCode = 20
		case "h:mm:ss":
			numFmtCode = 21
		case "m/d/yy h:mm":
			numFmtCode = 22
		case "@":
			numFmtCode = 50
		default:
			// If not a common format, log a warning and use General (0) or try to parse as int
			log.Printf("Warning: Unrecognized number format string '%s', using General (0).", numFmt)
			numFmtCode = 0
		}
		excelizeStyle.NumFmt = numFmtCode
	}

	// protection
	// if protectionData, ok := styleMap["protection"].(map[string]interface{}); ok {
	//     excelizeStyle.Locked = ...
	//     excelizeStyle.Hidden = ...
	//     // Not implemented yet, as it's not always used
	// }

	return excelizeStyle, nil
}

// getColorFromMap extracts a color string from a map[string]interface{}
func getColorFromMap(borderSideData map[string]interface{}) string {
	if colorData, ok := borderSideData["color"].(map[string]interface{}); ok {
		if rgb, ok := colorData["rgb"].(string); ok && len(rgb) == 6 {
			return rgb
		}
	}
	return "000000" // Default value if color is not found or invalid
}

// getStyleFromMap extracts a line style integer from a map[string]interface{}
func getStyleFromMap(borderSideData map[string]interface{}) int {
	if styleStr, ok := borderSideData["style"].(string); ok {
		// Map string values from openpyxl to excelize integer codes
		// See excelize documentation: https://pkg.go.dev/github.com/xuri/excelize/v2#FormatBorder
		switch styleStr {
		case "thin":
			return 2 // thin
		case "medium":
			return 6 // medium
		case "thick":
			return 8 // thick
			// Add other styles as needed
		default:
			return 0 // none
		}
	}
	return 0 // none as default
}

func main() {
	// Parse command-line arguments
	inputFile := flag.String("input", "", "Path to the input JSON file")
	outputFile := flag.String("output", "", "Path to the output XLSX file")
	flag.Parse()

	if *inputFile == "" || *outputFile == "" {
		fmt.Println("Usage: go_excel_exporter -input <input.json> -output <output.xlsx>")
		os.Exit(1)
	}

	// Read the JSON file
	jsonData, err := os.ReadFile(*inputFile)
	if err != nil {
		log.Fatalf("Error reading input file: %v", err)
	}

	// Parse JSON into Go structure
	var exportData ExportData

	err = json.Unmarshal(jsonData, &exportData)
	if err != nil {
		log.Fatalf("Error parsing JSON: %v", err)
	}

	// Create a new Excel file
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("Error closing file: %v", err)
		}
	}()

	// Process each sheet
	for i, sheet := range exportData.Sheets {
		var sheetName string
		if i == 0 {
			// Rename the first (default) sheet
			sheetName = sheet.Name
			if err := f.SetSheetName("Sheet1", sheetName); err != nil {
				log.Printf("Warning: could not rename default sheet to '%s': %v", sheetName, err)
				// If renaming fails, continue with sheet.Name, hoping SetSheetName creates it if it's not the first.
			}
		} else {
			// Create a new sheet for subsequent sheets
			_, err := f.NewSheet(sheet.Name) // <-- Fixed: ignore sheet index
			if err != nil {
				log.Printf("Warning: could not create new sheet '%s': %v", sheet.Name, err)
				continue // Skip this sheet if creation fails
			}
			sheetName = sheet.Name
			// Ensure the active sheet remains the first or last created?
			// f.SetActiveSheet(index) // Optional
		}

		// Populate data
		for rowIndex, row := range sheet.Data {
			for colIndex, cellValue := range row {
				// Excelize uses 1-based indexing
				cellRow := rowIndex + 1
				cellCol := colIndex + 1
				// Convert column number to name (A, B, ..., Z, AA, AB, ...)
				cellName, err := excelize.ColumnNumberToName(cellCol)
				if err != nil {
					log.Printf("Error converting column number %d to name: %v", cellCol, err)
					continue
				}
				cellAddress := fmt.Sprintf("%s%d", cellName, cellRow)

				if cellValue != nil {
					// Set cell value
					// f.SetCellValue(sheetName, cellAddress, *cellValue) // This method also works
					// Use a more specific method if the type is known, but SetCellValue is fine for general cases.
					if err := f.SetCellValue(sheetName, cellAddress, *cellValue); err != nil {
						log.Printf("Warning: could not set cell value at %s on sheet '%s': %v", cellAddress, sheetName, err)
					}
				}
			}
		}

		// Add formulas
		for _, formula := range sheet.Formulas {
			if err := f.SetCellFormula(sheetName, formula.Cell, formula.Formula); err != nil {
				log.Printf("Warning: could not set formula at %s on sheet '%s': %v", formula.Cell, sheetName, err)
			}
		}

		// --- START OF STYLE PROCESSING ---
		log.Printf("Processing %d styles for sheet '%s'", len(sheet.Styles), sheetName)
		for _, styleObj := range sheet.Styles {
			// 1. Convert JSON style definition to excelize.Style structure
			excelizeStyle, err := convertStyleToExcelizeOptions(styleObj.Style)
			if err != nil {
				log.Printf("Warning: could not convert style for range '%s' on sheet '%s': %v", styleObj.Range, sheetName, err)
				continue
			}

			// 2. Create the style in excelize
			styleID, err := f.NewStyle(excelizeStyle)
			if err != nil {
				log.Printf("Warning: could not create style in excelize for range '%s' on sheet '%s': %v", styleObj.Range, sheetName, err)
				continue
			}

			// 3. Apply the style to the range
			if err := f.SetCellStyle(sheetName, styleObj.Range, styleObj.Range, styleID); err != nil {
				log.Printf("Warning: could not apply style to range '%s' on sheet '%s': %v", styleObj.Range, sheetName, err)
				continue
			}
		}
		log.Printf("Finished processing styles for sheet '%s'", sheetName)
		// --- END OF STYLE PROCESSING ---

		// Add charts
		for _, chart := range sheet.Charts {
			// Create chart configuration
			chartConfig := &excelize.Chart{
				Type: convertChartType(chart.Type), // <-- Use our function
				// Series will be populated below
				Series: []excelize.ChartSeries{},
				// Title now takes []excelize.RichTextRun
				Title: []excelize.RichTextRun{{Text: chart.Title}},
			}

			// Populate data series for the chart
			for _, series := range chart.Series {
				chartConfig.Series = append(chartConfig.Series, excelize.ChartSeries{
					Name:       series.Name,
					Categories: series.Categories,
					Values:     series.Values,
				})
			}

			// Add the chart to the sheet
			if err := f.AddChart(sheetName, chart.Position, chartConfig); err != nil {
				log.Printf("Warning: could not add chart at %s on sheet '%s': %v", chart.Position, sheetName, err)
			}
		}
		// End of chart processing for the current sheet

		// Apply merged cells
		// f.MergeCell requires 4 arguments: sheet, coordinate for top-left cell, coordinate for bottom-right cell
		// The JSON contains a string like "A1:B2". We need to split this.
		for _, mergedCellRange := range sheet.MergedCells {
			// Split the range string "A1:B2" into "A1" and "B2"
			// This is a simple split, assumes no spaces and correct format.
			coords := splitRange(mergedCellRange)
			if len(coords) != 2 {
				log.Printf("Warning: Invalid merged cell range format '%s' on sheet '%s', skipping.", mergedCellRange, sheetName)
				continue
			}
			// coords[0] is top-left, coords[1] is bottom-right
			if err := f.MergeCell(sheetName, coords[0], coords[1]); err != nil {
				log.Printf("Warning: could not merge cells '%s' (from '%s' to '%s') on sheet '%s': %v", mergedCellRange, coords[0], coords[1], sheetName, err)
				continue
			}
		}
		// End of merged cell application for the current sheet

		// TODO: Process additional elements (images, tables, etc.)
	}
	// End of processing all sheets

	// Save the file
	if err := f.SaveAs(*outputFile); err != nil {
		log.Fatalf("Error saving file: %v", err)
	}

	fmt.Printf("Successfully exported to %s\n", *outputFile)
}

// splitRange splits a string like "A1:B2" into ["A1", "B2"]
func splitRange(rangeStr string) []string {
	// Find the colon separator
	colonIndex := -1
	for i, r := range rangeStr {
		if r == ':' {
			colonIndex = i
			break
		}
	}
	if colonIndex == -1 {
		// If no colon, return the whole string as a single element or an empty slice
		// For merged cells, this is invalid, so return empty.
		return []string{}
	}
	// Split the string
	topLeft := rangeStr[:colonIndex]
	bottomRight := rangeStr[colonIndex+1:]
	return []string{topLeft, bottomRight}
}
