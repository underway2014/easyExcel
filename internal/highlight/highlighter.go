package highlight

import (
	"easyExcel/internal/config"
	"easyExcel/internal/search"
	"easyExcel/pkg/excel"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// Highlighter handles cell and row highlighting
type Highlighter struct {
	wb          *excel.Workbook
	cfg         *config.HighlightConfig
	cellStyleID int
	rowStyleID  int
}

// NewHighlighter creates a new Highlighter
func NewHighlighter(wb *excel.Workbook, cfg *config.HighlightConfig) (*Highlighter, error) {
	h := &Highlighter{
		wb:  wb,
		cfg: cfg,
	}

	// Create cell highlight style only if color is not empty
	if cfg.Cell != "" {
		cellStyle, err := h.createStyle(cfg.Cell)
		if err != nil {
			return nil, fmt.Errorf("failed to create cell style: %w", err)
		}
		h.cellStyleID = cellStyle
	}

	// Create row highlight style only if color is not empty
	if cfg.Row != "" {
		rowStyle, err := h.createStyle(cfg.Row)
		if err != nil {
			return nil, fmt.Errorf("failed to create row style: %w", err)
		}
		h.rowStyleID = rowStyle
	}

	return h, nil
}

// createStyle creates a fill style for highlighting
func (h *Highlighter) createStyle(color string) (int, error) {
	if color == "" {
		return 0, nil
	}

	r, g, b, err := config.ParseColor(color)
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

// HighlightCells highlights only the matched cells
func (h *Highlighter) HighlightCells(results []search.MatchResult) error {
	// Skip if cell style was not created (color was empty)
	if h.cellStyleID == 0 {
		return nil
	}

	for _, result := range results {
		if err := h.wb.SetCellStyle(result.SheetName, result.Cell, h.cellStyleID); err != nil {
			return fmt.Errorf("failed to highlight cell %s: %w", result.Cell, err)
		}
	}
	return nil
}

// HighlightRows highlights the entire rows containing matches
func (h *Highlighter) HighlightRows(results []search.MatchResult, sheet string) error {
	// Skip if row style was not created (color was empty)
	if h.rowStyleID == 0 {
		return nil
	}

	// Get unique rows
	rowSet := make(map[int]bool)
	for _, result := range results {
		rowSet[result.Row] = true
	}

	maxCol, err := h.wb.GetMaxCol(sheet)
	if err != nil {
		return fmt.Errorf("failed to get max column: %w", err)
	}

	for row := range rowSet {
		if err := h.wb.SetRowStyle(sheet, row, 1, maxCol, h.rowStyleID); err != nil {
			return fmt.Errorf("failed to highlight row %d: %w", row, err)
		}
	}

	return nil
}

// HighlightMixed highlights cells with cell style and rows with row style
func (h *Highlighter) HighlightMixed(results []search.MatchResult, sheet string, highlightRow bool) error {
	// First highlight cells
	if err := h.HighlightCells(results); err != nil {
		return err
	}

	// Then highlight rows if requested
	if highlightRow {
		if err := h.HighlightRows(results, sheet); err != nil {
			return err
		}
	}

	return nil
}

// GetCellStyleID returns the cell highlight style ID
func (h *Highlighter) GetCellStyleID() int {
	return h.cellStyleID
}

// GetRowStyleID returns the row highlight style ID
func (h *Highlighter) GetRowStyleID() int {
	return h.rowStyleID
}
