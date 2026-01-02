package reader

import (
	"excel-lite/pkg/models"
)

// RowParser extracts structured row data from a table
type RowParser struct {
	config models.DetectionConfig
}

// NewRowParser creates a new row parser
func NewRowParser(config models.DetectionConfig) *RowParser {
	return &RowParser{config: config}
}

// NewDefaultRowParser creates a row parser with default config
func NewDefaultRowParser() *RowParser {
	return NewRowParser(models.DefaultConfig())
}

// ParseRows extracts rows from a grid given headers and boundary
func (rp *RowParser) ParseRows(grid [][]models.Cell, headers []string, headerRow int, boundary models.TableBoundary) []models.Row {
	var rows []models.Row

	// Start from the row after the header
	for rowIdx := headerRow + 1; rowIdx <= boundary.EndRow && rowIdx < len(grid); rowIdx++ {
		row := rp.parseRow(grid, rowIdx, headers, boundary)
		if row != nil && !rp.isEmptyRow(row) {
			rows = append(rows, *row)
		}
	}

	return rows
}

// parseRow converts a grid row to a structured Row
func (rp *RowParser) parseRow(grid [][]models.Cell, rowIdx int, headers []string, boundary models.TableBoundary) *models.Row {
	if rowIdx >= len(grid) {
		return nil
	}

	row := &models.Row{
		Index:  rowIdx,
		Values: make(map[string]models.Cell),
		Cells:  make([]models.Cell, 0, len(headers)),
	}

	for i, header := range headers {
		colIdx := boundary.StartCol + i
		var cell models.Cell

		if colIdx < len(grid[rowIdx]) {
			cell = grid[rowIdx][colIdx]
		} else {
			cell = models.Cell{
				Type:     models.CellTypeEmpty,
				Row:      rowIdx,
				Col:      colIdx,
				RawValue: "",
			}
		}

		row.Values[header] = cell
		row.Cells = append(row.Cells, cell)
	}

	return row
}

// isEmptyRow checks if all cells in the row are empty
func (rp *RowParser) isEmptyRow(row *models.Row) bool {
	if row == nil || len(row.Cells) == 0 {
		return true
	}
	for _, cell := range row.Cells {
		if !cell.IsEmpty() {
			return false
		}
	}
	return true
}

// ParseTable combines all components to parse a complete table
func (rp *RowParser) ParseTable(grid [][]models.Cell, boundary models.TableBoundary, headers []string, headerRow int, tableName string) models.Table {
	rows := rp.ParseRows(grid, headers, headerRow, boundary)

	return models.Table{
		Name:      tableName,
		Headers:   headers,
		Rows:      rows,
		StartRow:  boundary.StartRow,
		EndRow:    boundary.EndRow,
		StartCol:  boundary.StartCol,
		EndCol:    boundary.EndCol,
		HeaderRow: headerRow,
	}
}

// FilterRows filters rows based on a predicate function
func FilterRows(rows []models.Row, predicate func(models.Row) bool) []models.Row {
	var filtered []models.Row
	for _, row := range rows {
		if predicate(row) {
			filtered = append(filtered, row)
		}
	}
	return filtered
}

// MapRows transforms rows using a mapper function
func MapRows[T any](rows []models.Row, mapper func(models.Row) T) []T {
	result := make([]T, len(rows))
	for i, row := range rows {
		result[i] = mapper(row)
	}
	return result
}

// GetColumnValues extracts all values from a specific column
func GetColumnValues(rows []models.Row, header string) []models.Cell {
	values := make([]models.Cell, 0, len(rows))
	for _, row := range rows {
		if cell, ok := row.Get(header); ok {
			values = append(values, cell)
		}
	}
	return values
}
