package export

import (
	"easyExcel/internal/search"
	"easyExcel/pkg/excel"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// Exporter exports matched rows to a new Excel file
type Exporter struct {
	wb *excel.Workbook
}

// NewExporter creates a new Exporter
func NewExporter(wb *excel.Workbook) *Exporter {
	return &Exporter{wb: wb}
}

// ExportRows exports the matched rows to a new file
func (e *Exporter) ExportRows(results []search.MatchResult, outputPath string, highlightCfg *HighlightConfig) error {
	// Get unique rows
	rowSet := make(map[int]bool)
	sheetName := ""
	for _, result := range results {
		rowSet[result.Row] = true
		sheetName = result.SheetName
	}

	if sheetName == "" {
		return fmt.Errorf("no results to export")
	}

	// Create new workbook
	newWb := excel.NewEmptyWorkbook()
	newSheet := "Sheet1"
	newWb.NewSheet(newSheet)

	// Get source data
	maxCol, err := e.wb.GetMaxCol(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get max column: %w", err)
	}

	// Copy header row and matched rows
	newRow := 1
	for row := range rowSet {
		for col := 1; col <= maxCol; col++ {
			srcCell, _ := excelize.CoordinatesToCellName(col, row)
			dstCell, _ := excelize.CoordinatesToCellName(col, newRow)

			value, err := e.wb.GetCellValue(sheetName, srcCell)
			if err != nil {
				continue
			}

			if err := newWb.SetCellValue(newSheet, dstCell, value); err != nil {
				return fmt.Errorf("failed to set cell value: %w", err)
			}
		}
		newRow++
	}

	// Apply highlighting if configured
	if highlightCfg != nil {
		hl := &HighlightHelper{
			wb: newWb,
		}

		cellStyle, _ := hl.CreateStyle(highlightCfg.Cell)
		rowStyle, _ := hl.CreateStyle(highlightCfg.Row)

		// Highlight cells
		rowNum := 1
		for range rowSet {
			for col := 1; col <= maxCol; col++ {
				cell, _ := excelize.CoordinatesToCellName(col, rowNum)
				newWb.SetCellStyle(newSheet, cell, cellStyle)
			}
			if highlightCfg.Row != "" {
				newWb.SetRowStyle(newSheet, rowNum, 1, maxCol, rowStyle)
			}
			rowNum++
		}
	}

	// Save new workbook
	return newWb.Save(outputPath)
}

// ExportAllMatchedRows exports all matched rows including all columns
func (e *Exporter) ExportAllMatchedRows(sheetName string, rows []int, outputPath string, highlightCfg *HighlightConfig) error {
	if len(rows) == 0 {
		return fmt.Errorf("no rows to export")
	}

	// Create new workbook
	newWb := excel.NewEmptyWorkbook()
	newSheet := "Sheet1"
	newWb.NewSheet(newSheet)

	// Get source data
	maxCol, err := e.wb.GetMaxCol(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get max column: %w", err)
	}

	// Copy header (row 1) and matched rows
	destRow := 1

	// Copy header
	for col := 1; col <= maxCol; col++ {
		srcCell, _ := excelize.CoordinatesToCellName(col, 1)
		dstCell, _ := excelize.CoordinatesToCellName(col, destRow)

		value, err := e.wb.GetCellValue(sheetName, srcCell)
		if err != nil {
			continue
		}

		if err := newWb.SetCellValue(newSheet, dstCell, value); err != nil {
			return fmt.Errorf("failed to set header value: %w", err)
		}
	}
	destRow++

	// Copy matched rows
	for _, row := range rows {
		for col := 1; col <= maxCol; col++ {
			srcCell, _ := excelize.CoordinatesToCellName(col, row)
			dstCell, _ := excelize.CoordinatesToCellName(col, destRow)

			value, err := e.wb.GetCellValue(sheetName, srcCell)
			if err != nil {
				continue
			}

			if err := newWb.SetCellValue(newSheet, dstCell, value); err != nil {
				return fmt.Errorf("failed to set cell value: %w", err)
			}
		}
		destRow++
	}

	// Apply highlighting
	if highlightCfg != nil {
		hl := &HighlightHelper{wb: newWb}

		cellStyle, _ := hl.CreateStyle(highlightCfg.Cell)
		rowStyle, _ := hl.CreateStyle(highlightCfg.Row)

		// Highlight data rows (starting from row 2)
		for i := 0; i < len(rows); i++ {
			rowNum := i + 2
			for col := 1; col <= maxCol; col++ {
				cell, _ := excelize.CoordinatesToCellName(col, rowNum)
				newWb.SetCellStyle(newSheet, cell, cellStyle)
			}
			if highlightCfg.Row != "" {
				newWb.SetRowStyle(newSheet, rowNum, 1, maxCol, rowStyle)
			}
		}
	}

	return newWb.Save(outputPath)
}

// HighlightConfig holds highlight settings for export
type HighlightConfig struct {
	Cell string
	Row  string
}

// HighlightHelper helps with highlighting
type HighlightHelper struct {
	wb *excel.Workbook
}

// CreateStyle creates a highlight style
func (h *HighlightHelper) CreateStyle(color string) (int, error) {
	if color == "" {
		return 0, nil
	}

	r, g, b, err := parseColor(color)
	if err != nil {
		return 0, err
	}

	style := &excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{fmt.Sprintf("%02X%02X%02X", r, g, b)},
			Pattern: 1,
		},
	}

	return h.wb.NewStyle(style)
}

// parseColor parses a hex color string
func parseColor(color string) (int, int, int, error) {
	if len(color) == 0 {
		return 0, 0, 0, nil
	}

	color = fmt.Sprintf("%06s", color)
	if color[0] == '#' {
		color = color[1:]
	}

	if len(color) != 6 {
		return 0, 0, 0, fmt.Errorf("invalid color format")
	}

	var r, g, b int
	_, err := fmt.Sscanf(color, "%02X%02X%02X", &r, &g, &b)
	if err != nil {
		return 0, 0, 0, err
	}

	return r, g, b, nil
}
