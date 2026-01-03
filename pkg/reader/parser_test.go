package reader

import (
	"testing"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// =============================================================================
// Constructor Tests
// =============================================================================

func TestNewRowParser(t *testing.T) {
	config := models.DetectionConfig{
		MinColumns: 5,
		MinRows:    10,
	}

	rp := NewRowParser(config)

	if rp == nil {
		t.Fatal("NewRowParser() returned nil")
	}

	if rp.config.MinColumns != 5 {
		t.Errorf("config.MinColumns = %d, want 5", rp.config.MinColumns)
	}
}

func TestNewDefaultRowParser(t *testing.T) {
	rp := NewDefaultRowParser()

	if rp == nil {
		t.Fatal("NewDefaultRowParser() returned nil")
	}

	defaultConfig := models.DefaultConfig()
	if rp.config.MinRows != defaultConfig.MinRows {
		t.Errorf("config.MinRows = %d, want %d", rp.config.MinRows, defaultConfig.MinRows)
	}
}

// =============================================================================
// ParseRows Tests
// =============================================================================

func TestRowParser_ParseRows_Simple(t *testing.T) {
	rp := NewDefaultRowParser()

	grid := [][]models.Cell{
		{makeCell("Name", models.CellTypeString), makeCell("Age", models.CellTypeString)},
		{makeCell("Alice", models.CellTypeString), makeCell("30", models.CellTypeNumber)},
		{makeCell("Bob", models.CellTypeString), makeCell("25", models.CellTypeNumber)},
	}

	headers := []string{"Name", "Age"}
	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 1}

	rows := rp.ParseRows(grid, headers, 0, boundary)

	if len(rows) != 2 {
		t.Fatalf("ParseRows() returned %d rows, want 2", len(rows))
	}

	// Check first data row
	if rows[0].Index != 1 {
		t.Errorf("rows[0].Index = %d, want 1", rows[0].Index)
	}

	nameCell, ok := rows[0].Get("Name")
	if !ok {
		t.Error("rows[0].Get('Name') returned false")
	}
	if nameCell.AsString() != "Alice" {
		t.Errorf("rows[0]['Name'] = %q, want Alice", nameCell.AsString())
	}
}

func TestRowParser_ParseRows_EmptyGrid(t *testing.T) {
	rp := NewDefaultRowParser()

	rows := rp.ParseRows([][]models.Cell{}, []string{"A", "B"}, 0, models.TableBoundary{})

	if len(rows) != 0 {
		t.Errorf("ParseRows() with empty grid returned %d rows, want 0", len(rows))
	}
}

func TestRowParser_ParseRows_HeaderOnly(t *testing.T) {
	rp := NewDefaultRowParser()

	grid := [][]models.Cell{
		{makeCell("Name", models.CellTypeString), makeCell("Age", models.CellTypeString)},
	}

	headers := []string{"Name", "Age"}
	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 1}

	rows := rp.ParseRows(grid, headers, 0, boundary)

	if len(rows) != 0 {
		t.Errorf("ParseRows() with header only returned %d rows, want 0", len(rows))
	}
}

func TestRowParser_ParseRows_SkipsEmptyRows(t *testing.T) {
	rp := NewDefaultRowParser()

	grid := [][]models.Cell{
		{makeCell("Name", models.CellTypeString), makeCell("Age", models.CellTypeString)},
		{makeCell("Alice", models.CellTypeString), makeCell("30", models.CellTypeNumber)},
		{makeEmptyCell(), makeEmptyCell()}, // Empty row
		{makeCell("Bob", models.CellTypeString), makeCell("25", models.CellTypeNumber)},
	}

	headers := []string{"Name", "Age"}
	boundary := models.TableBoundary{StartRow: 0, EndRow: 3, StartCol: 0, EndCol: 1}

	rows := rp.ParseRows(grid, headers, 0, boundary)

	// Should have 2 rows (empty row skipped)
	if len(rows) != 2 {
		t.Errorf("ParseRows() returned %d rows, want 2 (empty skipped)", len(rows))
	}
}

func TestRowParser_ParseRows_PartialRows(t *testing.T) {
	rp := NewDefaultRowParser()

	grid := [][]models.Cell{
		{makeCell("A", models.CellTypeString), makeCell("B", models.CellTypeString), makeCell("C", models.CellTypeString)},
		{makeCell("1", models.CellTypeNumber), makeCell("2", models.CellTypeNumber)}, // Missing column C
		{makeCell("3", models.CellTypeNumber)},                                        // Missing columns B, C
	}

	headers := []string{"A", "B", "C"}
	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2}

	rows := rp.ParseRows(grid, headers, 0, boundary)

	if len(rows) != 2 {
		t.Fatalf("ParseRows() returned %d rows, want 2", len(rows))
	}

	// Check that missing cells are handled
	cCell, ok := rows[0].Get("C")
	if !ok {
		t.Error("rows[0].Get('C') returned false for partial row")
	}
	if !cCell.IsEmpty() {
		t.Error("Expected empty cell for missing column")
	}
}

func TestRowParser_ParseRows_OffsetBoundary(t *testing.T) {
	rp := NewDefaultRowParser()

	// Grid with table starting at column 2
	grid := [][]models.Cell{
		{makeEmptyCell(), makeEmptyCell(), makeCell("X", models.CellTypeString), makeCell("Y", models.CellTypeString)},
		{makeEmptyCell(), makeEmptyCell(), makeCell("1", models.CellTypeNumber), makeCell("2", models.CellTypeNumber)},
		{makeEmptyCell(), makeEmptyCell(), makeCell("3", models.CellTypeNumber), makeCell("4", models.CellTypeNumber)},
	}

	headers := []string{"X", "Y"}
	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 2, EndCol: 3}

	rows := rp.ParseRows(grid, headers, 0, boundary)

	if len(rows) != 2 {
		t.Fatalf("ParseRows() returned %d rows, want 2", len(rows))
	}

	xCell, _ := rows[0].Get("X")
	if xCell.AsString() != "1" {
		t.Errorf("rows[0]['X'] = %q, want 1", xCell.AsString())
	}
}

// =============================================================================
// ParseTable Tests
// =============================================================================

func TestRowParser_ParseTable(t *testing.T) {
	rp := NewDefaultRowParser()

	grid := [][]models.Cell{
		{makeCell("ID", models.CellTypeString), makeCell("Name", models.CellTypeString)},
		{makeCell("1", models.CellTypeNumber), makeCell("Alice", models.CellTypeString)},
		{makeCell("2", models.CellTypeNumber), makeCell("Bob", models.CellTypeString)},
	}

	headers := []string{"ID", "Name"}
	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 1}

	table := rp.ParseTable(grid, boundary, headers, 0, "TestTable")

	if table.Name != "TestTable" {
		t.Errorf("table.Name = %q, want TestTable", table.Name)
	}

	if len(table.Headers) != 2 {
		t.Errorf("len(table.Headers) = %d, want 2", len(table.Headers))
	}

	if len(table.Rows) != 2 {
		t.Errorf("len(table.Rows) = %d, want 2", len(table.Rows))
	}

	if table.StartRow != 0 || table.EndRow != 2 {
		t.Errorf("table row range = %d-%d, want 0-2", table.StartRow, table.EndRow)
	}

	if table.StartCol != 0 || table.EndCol != 1 {
		t.Errorf("table col range = %d-%d, want 0-1", table.StartCol, table.EndCol)
	}

	if table.HeaderRow != 0 {
		t.Errorf("table.HeaderRow = %d, want 0", table.HeaderRow)
	}
}

// =============================================================================
// isEmptyRow Tests
// =============================================================================

func TestRowParser_isEmptyRow(t *testing.T) {
	rp := NewDefaultRowParser()

	tests := []struct {
		name     string
		row      *models.Row
		expected bool
	}{
		{
			name:     "nil row",
			row:      nil,
			expected: true,
		},
		{
			name: "empty cells",
			row: &models.Row{
				Cells: []models.Cell{
					makeEmptyCell(),
					makeEmptyCell(),
				},
			},
			expected: true,
		},
		{
			name: "one non-empty",
			row: &models.Row{
				Cells: []models.Cell{
					makeEmptyCell(),
					makeCell("data", models.CellTypeString),
				},
			},
			expected: false,
		},
		{
			name: "all non-empty",
			row: &models.Row{
				Cells: []models.Cell{
					makeCell("a", models.CellTypeString),
					makeCell("b", models.CellTypeString),
				},
			},
			expected: false,
		},
		{
			name: "no cells",
			row: &models.Row{
				Cells: []models.Cell{},
			},
			expected: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := rp.isEmptyRow(tt.row)
			if got != tt.expected {
				t.Errorf("isEmptyRow() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// =============================================================================
// Helper Functions Tests
// =============================================================================

func TestFilterRows(t *testing.T) {
	// Create cells with proper float64 values
	rows := []models.Row{
		{Index: 1, Values: map[string]models.Cell{"Age": {Value: 20.0, Type: models.CellTypeNumber, RawValue: "20"}}},
		{Index: 2, Values: map[string]models.Cell{"Age": {Value: 30.0, Type: models.CellTypeNumber, RawValue: "30"}}},
		{Index: 3, Values: map[string]models.Cell{"Age": {Value: 25.0, Type: models.CellTypeNumber, RawValue: "25"}}},
		{Index: 4, Values: map[string]models.Cell{"Age": {Value: 35.0, Type: models.CellTypeNumber, RawValue: "35"}}},
	}

	// Filter rows with Age >= 25
	filtered := FilterRows(rows, func(row models.Row) bool {
		if cell, ok := row.Get("Age"); ok {
			if f, ok := cell.AsFloat(); ok {
				return f >= 25
			}
		}
		return false
	})

	if len(filtered) != 3 {
		t.Errorf("FilterRows() returned %d rows, want 3", len(filtered))
	}
}

func TestFilterRows_Empty(t *testing.T) {
	rows := []models.Row{}

	filtered := FilterRows(rows, func(row models.Row) bool {
		return true
	})

	if len(filtered) != 0 {
		t.Errorf("FilterRows() on empty slice returned %d rows, want 0", len(filtered))
	}
}

func TestFilterRows_NoMatch(t *testing.T) {
	rows := []models.Row{
		{Index: 1},
		{Index: 2},
	}

	filtered := FilterRows(rows, func(row models.Row) bool {
		return false
	})

	if len(filtered) != 0 {
		t.Errorf("FilterRows() with no matches returned %d rows, want 0", len(filtered))
	}
}

func TestMapRows(t *testing.T) {
	rows := []models.Row{
		{Index: 1, Values: map[string]models.Cell{"Name": makeCell("Alice", models.CellTypeString)}},
		{Index: 2, Values: map[string]models.Cell{"Name": makeCell("Bob", models.CellTypeString)}},
	}

	names := MapRows(rows, func(row models.Row) string {
		if cell, ok := row.Get("Name"); ok {
			return cell.AsString()
		}
		return ""
	})

	if len(names) != 2 {
		t.Fatalf("MapRows() returned %d items, want 2", len(names))
	}

	if names[0] != "Alice" {
		t.Errorf("names[0] = %q, want Alice", names[0])
	}

	if names[1] != "Bob" {
		t.Errorf("names[1] = %q, want Bob", names[1])
	}
}

func TestMapRows_Empty(t *testing.T) {
	rows := []models.Row{}

	result := MapRows(rows, func(row models.Row) int {
		return row.Index
	})

	if len(result) != 0 {
		t.Errorf("MapRows() on empty slice returned %d items, want 0", len(result))
	}
}

func TestGetColumnValues(t *testing.T) {
	rows := []models.Row{
		{Values: map[string]models.Cell{"A": makeCell("1", models.CellTypeNumber)}},
		{Values: map[string]models.Cell{"A": makeCell("2", models.CellTypeNumber)}},
		{Values: map[string]models.Cell{"A": makeCell("3", models.CellTypeNumber)}},
	}

	values := GetColumnValues(rows, "A")

	if len(values) != 3 {
		t.Fatalf("GetColumnValues() returned %d values, want 3", len(values))
	}

	for i, v := range values {
		expected := string(rune('1' + i))
		if v.AsString() != expected {
			t.Errorf("values[%d] = %q, want %q", i, v.AsString(), expected)
		}
	}
}

func TestGetColumnValues_MissingColumn(t *testing.T) {
	rows := []models.Row{
		{Values: map[string]models.Cell{"A": makeCell("1", models.CellTypeNumber)}},
		{Values: map[string]models.Cell{"B": makeCell("2", models.CellTypeNumber)}}, // No "A"
		{Values: map[string]models.Cell{"A": makeCell("3", models.CellTypeNumber)}},
	}

	values := GetColumnValues(rows, "A")

	// Should return 2 values (skips row without "A")
	if len(values) != 2 {
		t.Errorf("GetColumnValues() returned %d values, want 2", len(values))
	}
}

func TestGetColumnValues_Empty(t *testing.T) {
	values := GetColumnValues([]models.Row{}, "A")

	if len(values) != 0 {
		t.Errorf("GetColumnValues() on empty rows returned %d values, want 0", len(values))
	}
}

// =============================================================================
// Complex Scenarios
// =============================================================================

func TestRowParser_ParseRows_LargeDataset(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping large dataset test in short mode")
	}

	rp := NewDefaultRowParser()

	// Create 1000 rows x 10 columns
	numRows := 1000
	numCols := 10

	grid := make([][]models.Cell, numRows+1)

	// Header row
	grid[0] = make([]models.Cell, numCols)
	headers := make([]string, numCols)
	for c := 0; c < numCols; c++ {
		header := string(rune('A' + c))
		grid[0][c] = makeCell(header, models.CellTypeString)
		headers[c] = header
	}

	// Data rows
	for r := 1; r <= numRows; r++ {
		grid[r] = make([]models.Cell, numCols)
		for c := 0; c < numCols; c++ {
			grid[r][c] = makeCell("data", models.CellTypeString)
		}
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: numRows, StartCol: 0, EndCol: numCols - 1}

	rows := rp.ParseRows(grid, headers, 0, boundary)

	if len(rows) != numRows {
		t.Errorf("ParseRows() returned %d rows, want %d", len(rows), numRows)
	}
}

func TestRowParser_ParseRows_MixedDataTypes(t *testing.T) {
	rp := NewDefaultRowParser()

	grid := [][]models.Cell{
		{
			makeCell("String", models.CellTypeString),
			makeCell("Number", models.CellTypeString),
			makeCell("Bool", models.CellTypeString),
			makeCell("Date", models.CellTypeString),
		},
		{
			makeCell("hello", models.CellTypeString),
			makeCell("42.5", models.CellTypeNumber),
			makeCell("true", models.CellTypeBool),
			makeCell("2024-01-15", models.CellTypeDate),
		},
	}

	headers := []string{"String", "Number", "Bool", "Date"}
	boundary := models.TableBoundary{StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 3}

	rows := rp.ParseRows(grid, headers, 0, boundary)

	if len(rows) != 1 {
		t.Fatalf("ParseRows() returned %d rows, want 1", len(rows))
	}

	row := rows[0]

	// Check each type is preserved
	stringCell, _ := row.Get("String")
	if stringCell.Type != models.CellTypeString {
		t.Errorf("String cell type = %v, want CellTypeString", stringCell.Type)
	}

	numberCell, _ := row.Get("Number")
	if numberCell.Type != models.CellTypeNumber {
		t.Errorf("Number cell type = %v, want CellTypeNumber", numberCell.Type)
	}

	boolCell, _ := row.Get("Bool")
	if boolCell.Type != models.CellTypeBool {
		t.Errorf("Bool cell type = %v, want CellTypeBool", boolCell.Type)
	}
}
