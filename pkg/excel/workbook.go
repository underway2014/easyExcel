package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// CellPosition represents a cell position
type CellPosition struct {
	Sheet string
	Row   int
	Col   int
}

// Workbook wraps excelize.File
type Workbook struct {
	file   *excelize.File
	path   string
	modified bool
}

// NewWorkbook creates a new Workbook
func NewWorkbook(path string) (*Workbook, error) {
	file, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	return &Workbook{
		file:   file,
		path:   path,
		modified: false,
	}, nil
}

// NewEmptyWorkbook creates an empty workbook
func NewEmptyWorkbook() *Workbook {
	return &Workbook{
		file:   excelize.NewFile(),
		path:   "",
		modified: true,
	}
}

// GetSheets returns all sheet names
func (w *Workbook) GetSheets() []string {
	return w.file.GetSheetList()
}

// GetSheet returns the default sheet name
func (w *Workbook) GetSheet() string {
	return w.file.GetSheetName(0)
}

// GetCellValue gets the value of a cell
func (w *Workbook) GetCellValue(sheet, cell string) (string, error) {
	return w.file.GetCellValue(sheet, cell)
}

// GetCellValueByCoords gets the value by row and column
func (w *Workbook) GetCellValueByCoords(sheet string, row, col int) (string, error) {
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return "", err
	}
	return w.file.GetCellValue(sheet, cell)
}

// GetRow returns all values in a row
func (w *Workbook) GetRow(sheet string, row int) ([]string, error) {
	cols, err := w.file.GetCols(sheet)
	if err != nil {
		return nil, err
	}

	result := make([]string, len(cols))
	for i, col := range cols {
		if row-1 < len(col) {
			result[i] = col[row-1]
		} else {
			result[i] = ""
		}
	}
	return result, nil
}

// GetMaxRow returns the maximum row number in a sheet
func (w *Workbook) GetMaxRow(sheet string) (int, error) {
	rows, err := w.file.GetRows(sheet)
	if err != nil {
		return 0, err
	}
	return len(rows), nil
}

// GetMaxCol returns the maximum column number in a sheet
func (w *Workbook) GetMaxCol(sheet string) (int, error) {
	cols, err := w.file.GetCols(sheet)
	if err != nil {
		return 0, err
	}
	return len(cols), nil
}

// SetCellValue sets the value of a cell
func (w *Workbook) SetCellValue(sheet, cell, value string) error {
	err := w.file.SetCellValue(sheet, cell, value)
	if err != nil {
		return err
	}
	w.modified = true
	return nil
}

// SetCellValueByCoords sets the value by row and column
func (w *Workbook) SetCellValueByCoords(sheet string, row, col int, value string) error {
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return w.SetCellValue(sheet, cell, value)
}

// SetCellStyle sets the style of a cell
func (w *Workbook) SetCellStyle(sheet, cell string, styleID int) error {
	err := w.file.SetCellStyle(sheet, cell, cell, styleID)
	if err != nil {
		return err
	}
	w.modified = true
	return nil
}

// SetRowStyle sets the style for all cells in a row
func (w *Workbook) SetRowStyle(sheet string, row, startCol, endCol, styleID int) error {
	for col := startCol; col <= endCol; col++ {
		cell, err := excelize.CoordinatesToCellName(col, row)
		if err != nil {
			return err
		}
		if err := w.file.SetCellStyle(sheet, cell, cell, styleID); err != nil {
			return err
		}
	}
	w.modified = true
	return nil
}

// NewStyle creates a new style and returns its ID
func (w *Workbook) NewStyle(style *excelize.Style) (int, error) {
	return w.file.NewStyle(style)
}

// Save saves the workbook to the specified path
func (w *Workbook) Save(path string) error {
	if path == "" {
		path = w.path
	}
	if path == "" {
		return fmt.Errorf("no path specified")
	}
	if err := w.file.SaveAs(path); err != nil {
		return err
	}
	w.path = path
	w.modified = false

	// Close and reopen to verify save was successful
	if err := w.file.Close(); err != nil {
		return err
	}

	// Reopen the saved file to continue working with it
	newWb, err := excelize.OpenFile(path)
	if err != nil {
		return fmt.Errorf("failed to reopen saved file: %w", err)
	}
	w.file = newWb
	return nil
}

// Close closes the workbook
func (w *Workbook) Close() error {
	return w.file.Close()
}

// IsModified returns whether the workbook has been modified
func (w *Workbook) IsModified() bool {
	return w.modified
}

// GetFile returns the underlying excelize.File
func (w *Workbook) GetFile() *excelize.File {
	return w.file
}

// CopySheet copies a sheet
func (w *Workbook) CopySheet(source, target int) error {
	return w.file.CopySheet(source, target)
}

// NewSheet creates a new sheet
func (w *Workbook) NewSheet(name string) int {
	index, _ := w.file.NewSheet(name)
	return index
}

// DeleteSheet deletes a sheet
func (w *Workbook) DeleteSheet(name string) error {
	return w.file.DeleteSheet(name)
}
