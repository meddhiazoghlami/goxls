package reader

import (
	"strconv"
	"strings"
	"time"

	"excel-lite/pkg/models"

	"github.com/xuri/excelize/v2"
)

// SheetProcessor handles reading and processing individual sheets
type SheetProcessor struct {
	file   *ExcelFile
	config models.DetectionConfig
}

// NewSheetProcessor creates a new sheet processor with default config
func NewSheetProcessor(file *ExcelFile) *SheetProcessor {
	return &SheetProcessor{
		file:   file,
		config: models.DefaultConfig(),
	}
}

// NewSheetProcessorWithConfig creates a new sheet processor with custom config
func NewSheetProcessorWithConfig(file *ExcelFile, config models.DetectionConfig) *SheetProcessor {
	return &SheetProcessor{
		file:   file,
		config: config,
	}
}

// ReadSheet reads all cells from a sheet into a 2D grid
func (sp *SheetProcessor) ReadSheet(sheetName string) ([][]models.Cell, error) {
	rows, err := sp.file.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	if len(rows) == 0 {
		return [][]models.Cell{}, nil
	}

	// Find the maximum column count
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	// Build the cell grid
	grid := make([][]models.Cell, len(rows))
	for rowIdx, row := range rows {
		grid[rowIdx] = make([]models.Cell, maxCols)
		for colIdx := 0; colIdx < maxCols; colIdx++ {
			var rawValue string
			if colIdx < len(row) {
				rawValue = row[colIdx]
			}

			cellRef, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
			cellType := sp.detectCellType(sheetName, cellRef, rawValue)
			value := sp.parseValue(rawValue, cellType)

			// Check for formula
			var formula string
			var hasFormula bool
			if cellType == models.CellTypeFormula {
				if f, err := sp.file.GetCellFormula(sheetName, cellRef); err == nil && f != "" {
					formula = f
					hasFormula = true
				}
			}

			grid[rowIdx][colIdx] = models.Cell{
				Value:      value,
				Type:       cellType,
				Row:        rowIdx,
				Col:        colIdx,
				RawValue:   rawValue,
				IsMerged:   false,
				MergeRange: nil,
				Formula:    formula,
				HasFormula: hasFormula,
			}
		}
	}

	// Apply merge cell processing if enabled
	if sp.config.ExpandMergedCells || sp.config.TrackMergeMetadata {
		if err := sp.applyMerges(sheetName, grid); err != nil {
			// Log warning but don't fail - merges are an enhancement
			// The grid is still valid without merge info
			_ = err
		}
	}

	// Apply comments
	if err := sp.applyComments(sheetName, grid); err != nil {
		// Log warning but don't fail - comments are an enhancement
		_ = err
	}

	// Apply hyperlinks
	if err := sp.applyHyperlinks(sheetName, grid); err != nil {
		// Log warning but don't fail - hyperlinks are an enhancement
		_ = err
	}

	return grid, nil
}

// applyComments fetches and applies comment information to the grid
func (sp *SheetProcessor) applyComments(sheetName string, grid [][]models.Cell) error {
	comments, err := sp.file.GetComments(sheetName)
	if err != nil {
		return err
	}

	if len(comments) == 0 {
		return nil
	}

	// Build a map of cell reference to comment
	commentMap := make(map[string]string)
	for _, c := range comments {
		commentMap[c.Cell] = c.Text
	}

	// Apply comments to cells
	for rowIdx := range grid {
		for colIdx := range grid[rowIdx] {
			cellRef, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
			if comment, ok := commentMap[cellRef]; ok {
				grid[rowIdx][colIdx].Comment = comment
				grid[rowIdx][colIdx].HasComment = true
			}
		}
	}

	return nil
}

// applyHyperlinks fetches and applies hyperlink information to the grid
func (sp *SheetProcessor) applyHyperlinks(sheetName string, grid [][]models.Cell) error {
	for rowIdx := range grid {
		for colIdx := range grid[rowIdx] {
			cellRef, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
			link, err := sp.file.GetCellHyperLink(sheetName, cellRef)
			if err != nil {
				// Skip cells that error - continue with others
				continue
			}
			if link != "" {
				grid[rowIdx][colIdx].Hyperlink = link
				grid[rowIdx][colIdx].HasHyperlink = true
			}
		}
	}
	return nil
}

// applyMerges fetches and applies merge cell information to the grid
func (sp *SheetProcessor) applyMerges(sheetName string, grid [][]models.Cell) error {
	mergeInfos, err := sp.file.GetMergeCells(sheetName)
	if err != nil {
		return err
	}

	if len(mergeInfos) == 0 {
		return nil
	}

	mp := NewMergeProcessor(sp.config)
	parsedMerges, err := mp.ParseMergeCells(mergeInfos)
	if err != nil {
		return err
	}

	mp.ApplyMerges(grid, parsedMerges)
	return nil
}

// detectCellType determines the type of a cell value
func (sp *SheetProcessor) detectCellType(sheetName, cellRef, value string) models.CellType {
	// Try to get cell type from excelize first - this is important for formulas
	// which may have empty values when not evaluated
	ct, err := sp.file.GetCellType(sheetName, cellRef)
	if err == nil {
		switch ct {
		case excelize.CellTypeFormula:
			return models.CellTypeFormula
		case excelize.CellTypeNumber:
			return models.CellTypeNumber
		case excelize.CellTypeBool:
			return models.CellTypeBool
		case excelize.CellTypeSharedString, excelize.CellTypeInlineString:
			// Check if it might be a date formatted as string
			if isDateLike(value) {
				return models.CellTypeDate
			}
			return models.CellTypeString
		}
	}

	// Check for empty after checking excelize type
	if value == "" {
		return models.CellTypeEmpty
	}

	// Fallback: infer type from value
	return inferType(value)
}

// inferType infers the cell type from the raw value
func inferType(value string) models.CellType {
	if value == "" {
		return models.CellTypeEmpty
	}

	// Check for boolean
	lower := strings.ToLower(value)
	if lower == "true" || lower == "false" {
		return models.CellTypeBool
	}

	// Check for number
	if _, err := strconv.ParseFloat(value, 64); err == nil {
		return models.CellTypeNumber
	}

	// Check for date patterns
	if isDateLike(value) {
		return models.CellTypeDate
	}

	return models.CellTypeString
}

// isDateLike checks if a string looks like a date
func isDateLike(value string) bool {
	dateFormats := []string{
		"2006-01-02",
		"01/02/2006",
		"02/01/2006",
		"2006/01/02",
		"Jan 2, 2006",
		"January 2, 2006",
		"02-Jan-2006",
		"2006-01-02 15:04:05",
		"01/02/2006 15:04:05",
	}

	for _, format := range dateFormats {
		if _, err := time.Parse(format, value); err == nil {
			return true
		}
	}
	return false
}

// parseValue converts a raw string value to the appropriate Go type
func (sp *SheetProcessor) parseValue(value string, cellType models.CellType) interface{} {
	switch cellType {
	case models.CellTypeEmpty:
		return nil
	case models.CellTypeNumber:
		if f, err := strconv.ParseFloat(value, 64); err == nil {
			return f
		}
		return value
	case models.CellTypeBool:
		return strings.ToLower(value) == "true"
	case models.CellTypeDate:
		return parseDate(value)
	default:
		return value
	}
}

// parseDate attempts to parse a date string
func parseDate(value string) interface{} {
	dateFormats := []string{
		"2006-01-02",
		"01/02/2006",
		"02/01/2006",
		"2006/01/02",
		"Jan 2, 2006",
		"January 2, 2006",
		"02-Jan-2006",
		"2006-01-02 15:04:05",
		"01/02/2006 15:04:05",
	}

	for _, format := range dateFormats {
		if t, err := time.Parse(format, value); err == nil {
			return t
		}
	}
	return value
}

// GetDimensions returns the row and column count for a sheet
func (sp *SheetProcessor) GetDimensions(sheetName string) (rows, cols int, err error) {
	grid, err := sp.ReadSheet(sheetName)
	if err != nil {
		return 0, 0, err
	}

	rows = len(grid)
	if rows > 0 {
		cols = len(grid[0])
	}
	return rows, cols, nil
}

// IsRowEmpty checks if a row is entirely empty
func IsRowEmpty(row []models.Cell) bool {
	for _, cell := range row {
		if !cell.IsEmpty() {
			return false
		}
	}
	return true
}

// CountNonEmptyCells counts non-empty cells in a row
func CountNonEmptyCells(row []models.Cell) int {
	count := 0
	for _, cell := range row {
		if !cell.IsEmpty() {
			count++
		}
	}
	return count
}
