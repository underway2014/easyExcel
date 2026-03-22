package search

import (
	"easyExcel/internal/config"
	"easyExcel/pkg/excel"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// MatchResult represents a search match result
type MatchResult struct {
	SheetName string
	Row       int
	Col       int
	Cell      string
	Value     string
	MatchedCondition int // Index of the matched condition
}

// Engine is the search engine
type Engine struct {
	wb      *excel.Workbook
	matcher *Matcher
	cfg     *config.Config
}

// NewEngine creates a new search engine
func NewEngine(wb *excel.Workbook, cfg *config.Config) (*Engine, error) {
	matcher, err := NewMatcher(cfg.Conditions, cfg.Logic)
	if err != nil {
		return nil, err
	}

	return &Engine{
		wb:      wb,
		matcher: matcher,
		cfg:     cfg,
	}, nil
}

// Search searches for matches in the workbook
func (e *Engine) Search() ([]MatchResult, error) {
	var results []MatchResult

	sheets := e.wb.GetSheets()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found")
	}

	sheet := e.wb.GetSheet()

	// Get the columns we need to check
	columns := make(map[string]bool)
	for _, cond := range e.cfg.Conditions {
		columns[strings.ToUpper(cond.Column)] = true
	}

	maxRow, err := e.wb.GetMaxRow(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get max row: %w", err)
	}

	maxCol, err := e.wb.GetMaxCol(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get max col: %w", err)
	}

	// Iterate through rows
	for row := 1; row <= maxRow; row++ {
		// Build values map for this row
		values := make(map[string]string)

		for col := 1; col <= maxCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				continue
			}
			value, err := e.wb.GetCellValue(sheet, cell)
			if err != nil {
				continue
			}

			// Store value by column letter
			colLetter, _ := excelize.ColumnNumberToName(col)
			values[colLetter] = value
		}

		// Check if row matches all conditions
		if e.matcher.Evaluate(values) {
			// Find which cell(s) matched - use map to avoid duplicates
			addedCells := make(map[string]bool)
			for _, cond := range e.cfg.Conditions {
				colLetter := strings.ToUpper(cond.Column)
				cell := fmt.Sprintf("%s%d", colLetter, row)
				value := values[colLetter]

				// Skip if already added this cell
				if addedCells[cell] {
					continue
				}

				// Check if this condition actually matches
				condIndex := -1
				for i, c := range e.cfg.Conditions {
					if c.Column == cond.Column && c.Type == cond.Type && c.Value == cond.Value {
						condIndex = i
						break
					}
				}

				if condIndex >= 0 && e.matcher.MatchCondition(value, condIndex) {
					colNum, _ := excelize.ColumnNameToNumber(colLetter)
					results = append(results, MatchResult{
						SheetName:         sheet,
						Row:               row,
						Col:               colNum,
						Cell:              cell,
						Value:             value,
						MatchedCondition:  condIndex,
					})
					addedCells[cell] = true
				}
			}
		}
	}

	return results, nil
}

// SearchInSheet searches in a specific sheet
func (e *Engine) SearchInSheet(sheetName string) ([]MatchResult, error) {
	sheets := e.wb.GetSheets()
	hasSheet := false
	for _, s := range sheets {
		if s == sheetName {
			hasSheet = true
			break
		}
	}
	if !hasSheet {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}

	var results []MatchResult

	maxRow, err := e.wb.GetMaxRow(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get max row: %w", err)
	}

	maxCol, err := e.wb.GetMaxCol(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get max col: %w", err)
	}

	for row := 1; row <= maxRow; row++ {
		values := make(map[string]string)

		for col := 1; col <= maxCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				continue
			}
			value, err := e.wb.GetCellValue(sheetName, cell)
			if err != nil {
				continue
			}

			colLetter, _ := excelize.ColumnNumberToName(col)
			values[colLetter] = value
		}

		if e.matcher.Evaluate(values) {
			for _, cond := range e.cfg.Conditions {
				colLetter := strings.ToUpper(cond.Column)
				cell := fmt.Sprintf("%s%d", colLetter, row)
				value := values[colLetter]

				condIndex := -1
				for i, c := range e.cfg.Conditions {
					if c.Column == cond.Column && c.Type == cond.Type && c.Value == cond.Value {
						condIndex = i
						break
					}
				}

				if condIndex >= 0 && e.matcher.MatchCondition(value, condIndex) {
					colNum, _ := excelize.ColumnNameToNumber(colLetter)
					results = append(results, MatchResult{
						SheetName:       sheetName,
						Row:             row,
						Col:             colNum,
						Cell:            cell,
						Value:           value,
						MatchedCondition: condIndex,
					})
				}
			}
		}
	}

	return results, nil
}

// GetMatchedRows returns unique row numbers that have matches
func (e *Engine) GetMatchedRows(results []MatchResult) []int {
	rowSet := make(map[int]bool)
	for _, r := range results {
		rowSet[r.Row] = true
	}

	rows := make([]int, 0, len(rowSet))
	for row := range rowSet {
		rows = append(rows, row)
	}

	// Sort rows
	for i := 0; i < len(rows)-1; i++ {
		for j := i + 1; j < len(rows); j++ {
			if rows[i] > rows[j] {
				rows[i], rows[j] = rows[j], rows[i]
			}
		}
	}

	return rows
}

// GetConfig returns the config
func (e *Engine) GetConfig() *config.Config {
	return e.cfg
}
