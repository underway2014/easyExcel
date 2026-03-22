package gui

import (
	"easyExcel/internal/config"
	"easyExcel/internal/export"
	"easyExcel/internal/highlight"
	"easyExcel/internal/search"
	"easyExcel/pkg/excel"
	"fmt"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
)

// ConditionRow represents a single condition row
type ConditionRow struct {
	columnSelect *widget.Select
	typeSelect   *widget.Select
	valueEntry   *widget.Entry
	removeBtn    *widget.Button
}

// App represents the main application
type App struct {
	window       fyne.Window
	filePath     string
	wb           *excel.Workbook
	searchEngine *search.Engine
	results      []search.MatchResult
	matchedRows  []int

	// UI components
	fileLabel       *widget.Label
	searchBtn       *widget.Button
	highlightBtn    *widget.Button
	exportBtn       *widget.Button
	resultLabel     *widget.Label
	resultText      *widget.TextGrid
	statusLabel     *widget.Label
	conditionsBox   *fyne.Container
	addCondBtn      *widget.Button

	// Condition rows
	conditionRows []ConditionRow

	// Highlight widgets
	cellColorEntry *widget.Entry
	rowColorEntry  *widget.Entry
	cellCheck      *widget.Check
	rowCheck       *widget.Check

	// Search config
	logicType    string
	highlightCfg *config.HighlightConfig
}

// NewApp creates a new application
func NewApp() *App {
	return &App{
		logicType: "AND",
		highlightCfg: &config.HighlightConfig{
			Cell: "#FFFF00",
			Row:  "", // Empty by default, user must check "高亮整行"
		},
	}
}

// Run starts the application
func (a *App) Run() {
	fyneApp := app.New()
	a.window = fyneApp.NewWindow("Excel 搜索与高亮工具")
	a.window.Resize(fyne.NewSize(900, 700))
	a.window.SetFixedSize(false)

	a.setupUI()

	a.window.ShowAndRun()
}

func (a *App) setupUI() {
	// File selection
	fileLabel := widget.NewLabel("选择文件:")
	a.fileLabel = widget.NewLabel("未选择文件")
	a.fileLabel.TextStyle = fyne.TextStyle{Monospace: true}

	browseBtn := widget.NewButton("浏览...", a.onBrowse)

	fileContainer := container.NewVBox(
		container.NewHBox(fileLabel, browseBtn),
		a.fileLabel,
	)

	// Condition section
	condLabel := widget.NewLabel("查询条件")
	a.conditionsBox = container.NewVBox()
	a.addCondBtn = widget.NewButton("+ 添加条件", a.onAddCondition)

	// Add first condition row
	a.addConditionRow()

	// Logic selection
	logicLabel := widget.NewLabel("逻辑:")
	andRadio := widget.NewRadioGroup([]string{"AND", "OR"}, func(value string) {
		a.logicType = value
	})
	andRadio.SetSelected("AND")

	logicContainer := container.NewHBox(logicLabel, andRadio)

	// Highlight section
	highlightLabel := widget.NewLabel("高亮设置")
	highlightBox := a.createHighlightBox()

	// Buttons
	a.searchBtn = widget.NewButtonWithIcon("开始搜索", theme.SearchIcon(), a.onSearch)
	a.highlightBtn = widget.NewButtonWithIcon("高亮并保存", theme.DocumentSaveIcon(), a.onHighlight)
	a.exportBtn = widget.NewButtonWithIcon("导出结果", theme.FileIcon(), a.onExport)

	a.highlightBtn.Disable()
	a.exportBtn.Disable()

	buttonContainer := container.NewHBox(a.searchBtn, a.highlightBtn, a.exportBtn)

	// Results
	a.resultLabel = widget.NewLabel("")
	a.resultText = widget.NewTextGrid()
	a.resultText.Hide()

	resultScroll := container.NewScroll(a.resultText)
	resultScroll.SetMinSize(fyne.NewSize(0, 200))

	resultBox := container.NewVBox(a.resultLabel, resultScroll)

	// Status
	a.statusLabel = widget.NewLabel("")
	a.statusLabel.Hide()

	// Main layout
	content := container.NewVBox(
		widget.NewLabel("Excel 搜索与高亮工具"),
		widget.NewSeparator(),
		fileContainer,
		widget.NewSeparator(),
		condLabel,
		a.conditionsBox,
		a.addCondBtn,
		widget.NewSeparator(),
		logicContainer,
		widget.NewSeparator(),
		highlightLabel,
		highlightBox,
		widget.NewSeparator(),
		buttonContainer,
		widget.NewSeparator(),
		resultBox,
		widget.NewSeparator(),
		a.statusLabel,
	)

	scrollContent := container.NewScroll(content)
	scrollContent.SetMinSize(fyne.NewSize(850, 650))

	a.window.SetContent(scrollContent)
}

func (a *App) createConditionRow() ConditionRow {
	row := ConditionRow{}

	row.columnSelect = widget.NewSelect([]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}, nil)
	row.columnSelect.SetSelected("A")

	row.typeSelect = widget.NewSelect([]string{"contains", "eq", "startswith", "endswith", "gt", "lt", "gte", "lte", "regex", "empty"}, nil)
	row.typeSelect.SetSelected("contains")

	row.valueEntry = widget.NewEntry()
	row.valueEntry.SetPlaceHolder("输入搜索值...")

	row.removeBtn = widget.NewButton("删除", func() {
		a.removeConditionRow(row)
	})

	return row
}

func (a *App) addConditionRow() {
	row := a.createConditionRow()
	a.conditionRows = append(a.conditionRows, row)

	// Create row container - use Border to make value entry expand
	left := container.NewHBox(
		widget.NewLabel("列:"),
		row.columnSelect,
		widget.NewLabel(" 类型:"),
		row.typeSelect,
		widget.NewLabel(" 值:"),
	)
	right := row.removeBtn

	rowContainer := container.NewBorder(nil, nil, left, right, row.valueEntry)

	a.conditionsBox.Add(rowContainer)
	a.conditionsBox.Refresh()
	a.window.Content().Refresh()
}

func (a *App) removeConditionRow(row ConditionRow) {
	// Find and remove from conditionRows
	for i, r := range a.conditionRows {
		if r == row {
			a.conditionRows = append(a.conditionRows[:i], a.conditionRows[i+1:]...)
			break
		}
	}

	// Rebuild conditions box
	a.conditionsBox.RemoveAll()
	for _, r := range a.conditionRows {
		left := container.NewHBox(
			widget.NewLabel("列:"),
			r.columnSelect,
			widget.NewLabel(" 类型:"),
			r.typeSelect,
			widget.NewLabel(" 值:"),
		)
		right := r.removeBtn
		rowContainer := container.NewBorder(nil, nil, left, right, r.valueEntry)
		a.conditionsBox.Add(rowContainer)
	}
	a.conditionsBox.Refresh()
	a.window.Content().Refresh()
}

func (a *App) onAddCondition() {
	a.addConditionRow()
}

func (a *App) createHighlightBox() fyne.CanvasObject {
	a.cellColorEntry = widget.NewEntry()
	a.cellColorEntry.SetText("#FFFF00")

	a.rowColorEntry = widget.NewEntry()
	a.rowColorEntry.SetText("#FF9900")

	a.cellCheck = widget.NewCheck("高亮单元格", func(checked bool) {
		if checked {
			a.highlightCfg.Cell = a.cellColorEntry.Text
		} else {
			a.highlightCfg.Cell = ""
		}
	})
	a.cellCheck.SetChecked(true)
	// Set initial value
	a.highlightCfg.Cell = a.cellColorEntry.Text

	a.rowCheck = widget.NewCheck("高亮整行", func(checked bool) {
		if checked {
			a.highlightCfg.Row = a.rowColorEntry.Text
		} else {
			a.highlightCfg.Row = ""
		}
	})
	// rowCheck 默认不勾选，不需要额外设置

	// Use grid layout for better control
	cellRow := container.NewGridWithColumns(4,
		widget.NewLabel("单元格颜色:"),
		a.cellColorEntry,
		a.cellCheck,
		widget.NewLabel(""),
	)
	rowRow := container.NewGridWithColumns(4,
		widget.NewLabel("整行颜色:"),
		a.rowColorEntry,
		a.rowCheck,
		widget.NewLabel(""),
	)

	return container.NewVBox(cellRow, rowRow)
}

func (a *App) getConditions() []config.Condition {
	var conditions []config.Condition
	for i, row := range a.conditionRows {
		cond := config.Condition{
			Column: row.columnSelect.Selected,
			Type:   config.ConditionType(row.typeSelect.Selected),
			Value:  row.valueEntry.Text,
		}
		// First condition doesn't need logic, subsequent use the selected logic
		if i > 0 {
			cond.Logic = a.logicType
		}
		conditions = append(conditions, cond)
	}
	return conditions
}

func (a *App) onBrowse() {
	dialog.ShowFileOpen(func(uri fyne.URIReadCloser, err error) {
		if err != nil {
			return
		}
		if uri == nil {
			return
		}

		path := uri.URI().Path()
		a.filePath = path

		// Show truncated path if too long
		displayPath := path
		if len(displayPath) > 60 {
			displayPath = "..." + displayPath[len(displayPath)-57:]
		}
		a.fileLabel.SetText(displayPath)
		a.fileLabel.Refresh()

		// Load workbook
		wb, err := excel.NewWorkbook(path)
		if err != nil {
			dialog.ShowError(err, a.window)
			return
		}
		a.wb = wb

		a.statusLabel.SetText(fmt.Sprintf("已加载文件: %s", path))
		a.statusLabel.Show()
	}, a.window)
}

func (a *App) onSearch() {
	if a.wb == nil {
		dialog.ShowInformation("提示", "请先选择文件", a.window)
		return
	}

	conditions := a.getConditions()
	if len(conditions) == 0 {
		dialog.ShowInformation("提示", "请添加至少一个查询条件", a.window)
		return
	}

	// Update highlight config from entries
	a.highlightCfg.Cell = a.cellColorEntry.Text
	a.highlightCfg.Row = a.rowColorEntry.Text

	cfg := &config.Config{
		Conditions: conditions,
		Logic:      a.logicType,
		Highlight:  *a.highlightCfg,
	}

	engine, err := search.NewEngine(a.wb, cfg)
	if err != nil {
		dialog.ShowError(err, a.window)
		return
	}

	results, err := engine.Search()
	if err != nil {
		dialog.ShowError(err, a.window)
		return
	}

	a.results = results
	a.matchedRows = engine.GetMatchedRows(results)
	a.searchEngine = engine

	a.resultLabel.SetText(fmt.Sprintf("找到 %d 条匹配结果", len(a.results)))
	a.resultLabel.Show()

	// Update result text grid with matched rows preview
	a.updateResultText()

	if len(results) > 0 {
		a.highlightBtn.Enable()
		a.exportBtn.Enable()
	}

	a.statusLabel.SetText(fmt.Sprintf("搜索完成，找到 %d 行匹配数据", len(a.matchedRows)))
}

func (a *App) updateResultText() {
	sheet := a.wb.GetSheet()

	grid := widget.NewTextGrid()
	grid.Show()

	// Build preview text
	text := fmt.Sprintf("找到 %d 行匹配结果:\n\n", len(a.matchedRows))

	for _, row := range a.matchedRows {
		rowData, _ := a.wb.GetRow(sheet, row)
		text += fmt.Sprintf("行 %d: %v\n", row, rowData)
		if len(text) > 5000 {
			text += "\n...(结果过多，已截断)"
			break
		}
	}

	grid.SetText(text)
	a.resultText = grid
	a.resultText.Show()
}

func (a *App) onHighlight() {
	if a.wb == nil || len(a.results) == 0 {
		return
	}

	// Update highlight config based on checkbox states (not just text values)
	cellChecked := a.cellCheck.Checked
	rowChecked := a.rowCheck.Checked

	// Update highlight config from current UI values
	if cellChecked {
		a.highlightCfg.Cell = a.cellColorEntry.Text
	} else {
		a.highlightCfg.Cell = ""
	}

	if rowChecked {
		a.highlightCfg.Row = a.rowColorEntry.Text
	} else {
		a.highlightCfg.Row = ""
	}

	// Check if any highlighting is enabled
	if a.highlightCfg.Cell == "" && a.highlightCfg.Row == "" {
		dialog.ShowInformation("提示", "请至少选择一种高亮方式", a.window)
		return
	}

	// Get output path
	dialog.ShowFileSave(func(uri fyne.URIWriteCloser, err error) {
		if err != nil {
			return
		}
		if uri == nil {
			return
		}

		outputPath := uri.URI().Path()

		hl, err := highlight.NewHighlighter(a.wb, a.highlightCfg)
		if err != nil {
			dialog.ShowError(err, a.window)
			return
		}

		sheet := a.wb.GetSheet()

		// Highlight cells only if enabled
		if a.highlightCfg.Cell != "" {
			if err := hl.HighlightCells(a.results); err != nil {
				dialog.ShowError(err, a.window)
				return
			}
		}

		// Highlight rows only if enabled
		if a.highlightCfg.Row != "" {
			if err := hl.HighlightRows(a.results, sheet); err != nil {
				dialog.ShowError(err, a.window)
				return
			}
		}

		// Save
		if err := a.wb.Save(outputPath); err != nil {
			dialog.ShowError(err, a.window)
			return
		}

		dialog.ShowInformation("完成", fmt.Sprintf("文件已保存到: %s", outputPath), a.window)
	}, a.window)
}

func (a *App) onExport() {
	if a.wb == nil || len(a.matchedRows) == 0 {
		return
	}

	// Reload the original file to get fresh data without highlights
	freshWb, err := excel.NewWorkbook(a.filePath)
	if err != nil {
		dialog.ShowError(err, a.window)
		return
	}

	dialog.ShowFileSave(func(uri fyne.URIWriteCloser, err error) {
		if err != nil {
			freshWb.Close()
			return
		}
		if uri == nil {
			freshWb.Close()
			return
		}

		outputPath := uri.URI().Path()

		exporter := export.NewExporter(freshWb)

		hlCfg := &export.HighlightConfig{
			Cell: a.highlightCfg.Cell,
			Row:  a.highlightCfg.Row,
		}

		sheet := freshWb.GetSheet()

		if err := exporter.ExportAllMatchedRows(sheet, a.matchedRows, outputPath, hlCfg); err != nil {
			freshWb.Close()
			dialog.ShowError(err, a.window)
			return
		}

		freshWb.Close()
		dialog.ShowInformation("完成", fmt.Sprintf("已导出 %d 行到: %s", len(a.matchedRows), outputPath), a.window)
	}, a.window)
}
