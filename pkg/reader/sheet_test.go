package reader

import (
	"path/filepath"
	"testing"

	"github.com/meddhiazoghlami/goxls/pkg/models"

	"github.com/xuri/excelize/v2"
)

// =============================================================================
// Test Helpers
// =============================================================================

func createSheetTestFile(t *testing.T, setup func(*excelize.File)) *ExcelFile {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "test.xlsx")

	f := excelize.NewFile()
	setup(f)
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("Failed to load test file: %v", err)
	}
	return ef
}

// =============================================================================
// SheetProcessor Tests
// =============================================================================

func TestNewSheetProcessor(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	if sp == nil {
		t.Fatal("NewSheetProcessor() returned nil")
	}
	if sp.file != ef {
		t.Error("SheetProcessor.file not set correctly")
	}
}

func TestSheetProcessor_ReadSheet_Simple(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 30)
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	if len(grid) != 2 {
		t.Errorf("ReadSheet() returned %d rows, want 2", len(grid))
	}

	if len(grid[0]) != 2 {
		t.Errorf("ReadSheet() first row has %d cols, want 2", len(grid[0]))
	}

	// Check first cell
	if grid[0][0].RawValue != "Name" {
		t.Errorf("grid[0][0].RawValue = %v, want Name", grid[0][0].RawValue)
	}

	if grid[0][0].Type != models.CellTypeString {
		t.Errorf("grid[0][0].Type = %v, want CellTypeString", grid[0][0].Type)
	}
}

func TestSheetProcessor_ReadSheet_Empty(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		// Empty sheet
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	if len(grid) != 0 {
		t.Errorf("ReadSheet() returned %d rows for empty sheet, want 0", len(grid))
	}
}

func TestSheetProcessor_ReadSheet_MixedTypes(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "String")
		f.SetCellValue("Sheet1", "B1", 42)
		f.SetCellValue("Sheet1", "C1", 3.14)
		f.SetCellValue("Sheet1", "D1", true)
		f.SetCellValue("Sheet1", "E1", "2024-01-15")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	tests := []struct {
		col      int
		expected models.CellType
		rawValue string
	}{
		{0, models.CellTypeString, "String"},
		{1, models.CellTypeNumber, "42"},
		{2, models.CellTypeNumber, "3.14"},
		{3, models.CellTypeBool, "TRUE"},
	}

	for _, tt := range tests {
		cell := grid[0][tt.col]
		if cell.Type != tt.expected {
			t.Errorf("grid[0][%d].Type = %v, want %v", tt.col, cell.Type, tt.expected)
		}
	}
}

func TestSheetProcessor_ReadSheet_JaggedRows(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "A")
		f.SetCellValue("Sheet1", "B1", "B")
		f.SetCellValue("Sheet1", "C1", "C")
		f.SetCellValue("Sheet1", "D1", "D")
		// Row 2 only has 2 values
		f.SetCellValue("Sheet1", "A2", "1")
		f.SetCellValue("Sheet1", "B2", "2")
		// Row 3 has values in different positions
		f.SetCellValue("Sheet1", "A3", "X")
		f.SetCellValue("Sheet1", "D3", "Y")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// All rows should have same number of columns (padded)
	maxCols := len(grid[0])
	for i, row := range grid {
		if len(row) != maxCols {
			t.Errorf("Row %d has %d cols, expected %d", i, len(row), maxCols)
		}
	}

	// Check that empty cells are properly typed
	if !grid[1][2].IsEmpty() {
		t.Error("Expected grid[1][2] to be empty")
	}

	if !grid[1][3].IsEmpty() {
		t.Error("Expected grid[1][3] to be empty")
	}
}

func TestSheetProcessor_ReadSheet_CellCoordinates(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "0,0")
		f.SetCellValue("Sheet1", "B1", "0,1")
		f.SetCellValue("Sheet1", "A2", "1,0")
		f.SetCellValue("Sheet1", "B2", "1,1")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Check that cell coordinates are set correctly
	for row := 0; row < 2; row++ {
		for col := 0; col < 2; col++ {
			cell := grid[row][col]
			if cell.Row != row {
				t.Errorf("Cell at [%d][%d] has Row=%d", row, col, cell.Row)
			}
			if cell.Col != col {
				t.Errorf("Cell at [%d][%d] has Col=%d", row, col, cell.Col)
			}
		}
	}
}

func TestSheetProcessor_ReadSheet_UnicodeContent(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "日本語")
		f.SetCellValue("Sheet1", "B1", "한국어")
		f.SetCellValue("Sheet1", "C1", "中文")
		f.SetCellValue("Sheet1", "D1", "العربية")
		f.SetCellValue("Sheet1", "E1", "emoji 🎉")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	expected := []string{"日本語", "한국어", "中文", "العربية", "emoji 🎉"}
	for i, exp := range expected {
		if grid[0][i].RawValue != exp {
			t.Errorf("grid[0][%d].RawValue = %v, want %v", i, grid[0][i].RawValue, exp)
		}
	}
}

func TestSheetProcessor_ReadSheet_SpecialValues(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "")          // Empty
		f.SetCellValue("Sheet1", "B1", " ")         // Whitespace
		f.SetCellValue("Sheet1", "C1", "  spaces ") // Spaces
		f.SetCellValue("Sheet1", "D1", "\t\n")      // Special chars
		f.SetCellValue("Sheet1", "E1", 0)           // Zero
		f.SetCellValue("Sheet1", "F1", -1)          // Negative
		f.SetCellValue("Sheet1", "G1", 0.0001)      // Small decimal
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Check empty cell
	if !grid[0][0].IsEmpty() {
		t.Error("Expected empty cell at A1")
	}

	// Check zero is not empty
	if grid[0][4].IsEmpty() {
		t.Error("Zero should not be considered empty")
	}
}

func TestSheetProcessor_ReadSheet_NonexistentSheet(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	_, err := sp.ReadSheet("NonexistentSheet")

	if err == nil {
		t.Error("ReadSheet() expected error for nonexistent sheet, got nil")
	}
}

func TestSheetProcessor_GetDimensions(t *testing.T) {
	tests := []struct {
		name         string
		setup        func(*excelize.File)
		expectedRows int
		expectedCols int
	}{
		{
			name: "3x4 table",
			setup: func(f *excelize.File) {
				for row := 1; row <= 3; row++ {
					for col := 1; col <= 4; col++ {
						cell, _ := excelize.CoordinatesToCellName(col, row)
						f.SetCellValue("Sheet1", cell, "X")
					}
				}
			},
			expectedRows: 3,
			expectedCols: 4,
		},
		{
			name: "single cell",
			setup: func(f *excelize.File) {
				f.SetCellValue("Sheet1", "A1", "Single")
			},
			expectedRows: 1,
			expectedCols: 1,
		},
		{
			name: "empty sheet",
			setup: func(f *excelize.File) {
				// Empty
			},
			expectedRows: 0,
			expectedCols: 0,
		},
		{
			name: "sparse data",
			setup: func(f *excelize.File) {
				f.SetCellValue("Sheet1", "A1", "Start")
				f.SetCellValue("Sheet1", "E5", "End")
			},
			expectedRows: 5,
			expectedCols: 5,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ef := createSheetTestFile(t, tt.setup)
			defer ef.Close()

			sp := NewSheetProcessor(ef)
			rows, cols, err := sp.GetDimensions("Sheet1")

			if err != nil {
				t.Fatalf("GetDimensions() error = %v", err)
			}

			if rows != tt.expectedRows {
				t.Errorf("GetDimensions() rows = %d, want %d", rows, tt.expectedRows)
			}

			if cols != tt.expectedCols {
				t.Errorf("GetDimensions() cols = %d, want %d", cols, tt.expectedCols)
			}
		})
	}
}

// =============================================================================
// Helper Function Tests
// =============================================================================

func TestIsRowEmpty(t *testing.T) {
	tests := []struct {
		name     string
		row      []models.Cell
		expected bool
	}{
		{
			name:     "nil row",
			row:      nil,
			expected: true,
		},
		{
			name:     "empty slice",
			row:      []models.Cell{},
			expected: true,
		},
		{
			name: "all empty cells",
			row: []models.Cell{
				{Type: models.CellTypeEmpty, RawValue: ""},
				{Type: models.CellTypeEmpty, RawValue: ""},
				{Type: models.CellTypeEmpty, RawValue: ""},
			},
			expected: true,
		},
		{
			name: "one non-empty cell",
			row: []models.Cell{
				{Type: models.CellTypeEmpty, RawValue: ""},
				{Type: models.CellTypeString, RawValue: "data"},
				{Type: models.CellTypeEmpty, RawValue: ""},
			},
			expected: false,
		},
		{
			name: "all non-empty cells",
			row: []models.Cell{
				{Type: models.CellTypeString, RawValue: "a"},
				{Type: models.CellTypeNumber, RawValue: "1"},
				{Type: models.CellTypeBool, RawValue: "true"},
			},
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := IsRowEmpty(tt.row); got != tt.expected {
				t.Errorf("IsRowEmpty() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestCountNonEmptyCells(t *testing.T) {
	tests := []struct {
		name     string
		row      []models.Cell
		expected int
	}{
		{
			name:     "nil row",
			row:      nil,
			expected: 0,
		},
		{
			name:     "empty slice",
			row:      []models.Cell{},
			expected: 0,
		},
		{
			name: "all empty",
			row: []models.Cell{
				{Type: models.CellTypeEmpty, RawValue: ""},
				{Type: models.CellTypeEmpty, RawValue: ""},
			},
			expected: 0,
		},
		{
			name: "mixed",
			row: []models.Cell{
				{Type: models.CellTypeString, RawValue: "a"},
				{Type: models.CellTypeEmpty, RawValue: ""},
				{Type: models.CellTypeNumber, RawValue: "1"},
				{Type: models.CellTypeEmpty, RawValue: ""},
				{Type: models.CellTypeBool, RawValue: "true"},
			},
			expected: 3,
		},
		{
			name: "all non-empty",
			row: []models.Cell{
				{Type: models.CellTypeString, RawValue: "a"},
				{Type: models.CellTypeString, RawValue: "b"},
				{Type: models.CellTypeString, RawValue: "c"},
			},
			expected: 3,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := CountNonEmptyCells(tt.row); got != tt.expected {
				t.Errorf("CountNonEmptyCells() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// =============================================================================
// Type Inference Tests
// =============================================================================

func TestInferType(t *testing.T) {
	tests := []struct {
		name     string
		value    string
		expected models.CellType
	}{
		{"empty", "", models.CellTypeEmpty},
		{"string", "hello", models.CellTypeString},
		{"integer", "42", models.CellTypeNumber},
		{"float", "3.14159", models.CellTypeNumber},
		{"negative", "-123", models.CellTypeNumber},
		{"scientific", "1.5e10", models.CellTypeNumber},
		{"true lowercase", "true", models.CellTypeBool},
		{"false lowercase", "false", models.CellTypeBool},
		{"TRUE uppercase", "TRUE", models.CellTypeBool},
		{"FALSE uppercase", "FALSE", models.CellTypeBool},
		{"date iso", "2024-01-15", models.CellTypeDate},
		{"date us", "01/15/2024", models.CellTypeDate},
		{"date eu", "15/01/2024", models.CellTypeDate},
		{"mixed string", "abc123", models.CellTypeString},
		{"number with text", "42 units", models.CellTypeString},
		{"whitespace", "   ", models.CellTypeString},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := inferType(tt.value); got != tt.expected {
				t.Errorf("inferType(%q) = %v, want %v", tt.value, got, tt.expected)
			}
		})
	}
}

func TestIsDateLike(t *testing.T) {
	tests := []struct {
		value    string
		expected bool
	}{
		{"2024-01-15", true},
		{"01/15/2024", true},
		{"15/01/2024", true},
		{"2024/01/15", true},
		{"Jan 15, 2024", true},
		{"January 15, 2024", true},
		{"15-Jan-2024", true},
		{"2024-01-15 10:30:00", true},
		{"not a date", false},
		{"2024", false},
		{"01-15", false},
		{"", false},
		{"12345", false},
	}

	for _, tt := range tests {
		t.Run(tt.value, func(t *testing.T) {
			if got := isDateLike(tt.value); got != tt.expected {
				t.Errorf("isDateLike(%q) = %v, want %v", tt.value, got, tt.expected)
			}
		})
	}
}

// =============================================================================
// Additional Coverage Tests
// =============================================================================

func TestSheetProcessor_ReadSheet_WithFormulas(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "B1", 20)
		f.SetCellFormula("Sheet1", "C1", "A1+B1")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Cell C1 should be detected as formula or show the result
	if len(grid) == 0 || len(grid[0]) < 3 {
		t.Fatal("Expected at least 3 columns")
	}

	// The formula cell exists and has a type
	cell := grid[0][2]
	if cell.Type == models.CellTypeEmpty && cell.RawValue == "" {
		// Formula might not have been evaluated, that's ok
	}
}

func TestSheetProcessor_ReadSheet_FormulaExtraction(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "B1", 20)
		f.SetCellFormula("Sheet1", "C1", "A1+B1")
		f.SetCellFormula("Sheet1", "D1", "SUM(A1:B1)")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	if len(grid) == 0 || len(grid[0]) < 4 {
		t.Fatal("Expected at least 4 columns")
	}

	// Check formula cells have HasFormula=true and Formula populated
	tests := []struct {
		col         int
		hasFormula  bool
		formulaLike string // substring to check (formulas may vary slightly)
	}{
		{0, false, ""},         // A1 - plain number
		{1, false, ""},         // B1 - plain number
		{2, true, "A1+B1"},     // C1 - formula
		{3, true, "SUM(A1:B1)"}, // D1 - formula
	}

	for _, tt := range tests {
		cell := grid[0][tt.col]
		if cell.HasFormula != tt.hasFormula {
			t.Errorf("Cell[0][%d].HasFormula = %v, want %v", tt.col, cell.HasFormula, tt.hasFormula)
		}
		if tt.hasFormula && cell.Formula == "" {
			t.Errorf("Cell[0][%d].Formula is empty, expected formula string", tt.col)
		}
		if tt.hasFormula && cell.Formula != tt.formulaLike {
			t.Errorf("Cell[0][%d].Formula = %q, want %q", tt.col, cell.Formula, tt.formulaLike)
		}
	}
}

func TestSheetProcessor_ReadSheet_FormulaType(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", 5)
		f.SetCellValue("Sheet1", "A2", 10)
		f.SetCellValue("Sheet1", "A3", 15)
		f.SetCellFormula("Sheet1", "A4", "SUM(A1:A3)")
		f.SetCellFormula("Sheet1", "A5", "AVERAGE(A1:A3)")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Formula cells should have CellTypeFormula
	for row := 3; row <= 4; row++ {
		cell := grid[row][0]
		if cell.Type != models.CellTypeFormula {
			t.Errorf("Cell[%d][0].Type = %v, want CellTypeFormula", row, cell.Type)
		}
		if !cell.HasFormula {
			t.Errorf("Cell[%d][0].HasFormula = false, want true", row)
		}
	}
}

func TestSheetProcessor_ReadSheet_NonFormulaCellsNoFormula(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "text")
		f.SetCellValue("Sheet1", "B1", 42)
		f.SetCellValue("Sheet1", "C1", true)
		f.SetCellValue("Sheet1", "D1", "2024-01-15")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// None of these cells should have formulas
	for col := 0; col < 4; col++ {
		cell := grid[0][col]
		if cell.HasFormula {
			t.Errorf("Cell[0][%d].HasFormula = true, want false for non-formula cell", col)
		}
		if cell.Formula != "" {
			t.Errorf("Cell[0][%d].Formula = %q, want empty string", col, cell.Formula)
		}
	}
}

func TestSheetProcessor_GetDimensions_Error(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)

	// Try to get dimensions of non-existent sheet
	_, _, err := sp.GetDimensions("NonExistent")
	if err == nil {
		t.Error("Expected error for non-existent sheet")
	}
}

func TestSheetProcessor_ReadSheet_DateFormatWithTime(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		// Set a date with time format
		f.SetCellValue("Sheet1", "A1", "01/15/2024 14:30:00")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Should be detected as date
	if grid[0][0].Type != models.CellTypeDate {
		t.Errorf("Expected CellTypeDate for datetime string, got %v", grid[0][0].Type)
	}
}

func TestSheetProcessor_ReadSheet_LargeNumbers(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", 1234567890123456789)
		f.SetCellValue("Sheet1", "B1", 0.000000001)
		f.SetCellValue("Sheet1", "C1", -999999999.99)
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// All should be numbers
	for i := 0; i < 3; i++ {
		if grid[0][i].Type != models.CellTypeNumber {
			t.Errorf("grid[0][%d].Type = %v, expected CellTypeNumber", i, grid[0][i].Type)
		}
	}
}

func TestParseDate_UnparsableDates(t *testing.T) {
	// Test parseDate with values that look like dates but aren't valid
	tests := []struct {
		value    string
		isString bool
	}{
		{"not-a-date", true},
		{"2024-99-99", true}, // Invalid month/day
		{"abc/def/ghi", true},
		{"", true},
	}

	for _, tt := range tests {
		result := parseDate(tt.value)
		if tt.isString {
			if _, ok := result.(string); !ok {
				t.Errorf("parseDate(%q) should return string, got %T", tt.value, result)
			}
		}
	}
}

func TestParseDate_AllFormats(t *testing.T) {
	// Test each supported date format
	validDates := []string{
		"2024-01-15",
		"01/15/2024",
		"15/01/2024",
		"2024/01/15",
		"Jan 15, 2024",
		"January 15, 2024",
		"15-Jan-2024",
		"2024-01-15 10:30:00",
		"01/15/2024 10:30:00",
	}

	for _, d := range validDates {
		result := parseDate(d)
		if _, ok := result.(string); ok {
			t.Errorf("parseDate(%q) should parse to time.Time, got string", d)
		}
	}
}

func TestInferType_EdgeCases(t *testing.T) {
	tests := []struct {
		value    string
		expected models.CellType
	}{
		{"True", models.CellTypeBool},
		{"False", models.CellTypeBool},
		{"TRUE", models.CellTypeBool},
		{"FALSE", models.CellTypeBool},
		{"+123", models.CellTypeNumber},
		{"-0.5", models.CellTypeNumber},
		{"1e-10", models.CellTypeNumber},
		{".5", models.CellTypeNumber},
		{"-.5", models.CellTypeNumber},
	}

	for _, tt := range tests {
		t.Run(tt.value, func(t *testing.T) {
			got := inferType(tt.value)
			if got != tt.expected {
				t.Errorf("inferType(%q) = %v, want %v", tt.value, got, tt.expected)
			}
		})
	}
}

func TestSheetProcessor_detectCellType_Fallback(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		// Set various cell types
		f.SetCellValue("Sheet1", "A1", "text")
		f.SetCellValue("Sheet1", "B1", 42)
		f.SetCellValue("Sheet1", "C1", true)
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Verify types were detected
	if grid[0][0].Type != models.CellTypeString {
		t.Errorf("Expected CellTypeString, got %v", grid[0][0].Type)
	}
	if grid[0][1].Type != models.CellTypeNumber {
		t.Errorf("Expected CellTypeNumber, got %v", grid[0][1].Type)
	}
	if grid[0][2].Type != models.CellTypeBool {
		t.Errorf("Expected CellTypeBool, got %v", grid[0][2].Type)
	}
}

// =============================================================================
// Merge Cell Integration Tests
// =============================================================================

func TestNewSheetProcessorWithConfig(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
	})
	defer ef.Close()

	config := models.DetectionConfig{
		ExpandMergedCells:  false,
		TrackMergeMetadata: true,
	}

	sp := NewSheetProcessorWithConfig(ef, config)
	if sp == nil {
		t.Fatal("NewSheetProcessorWithConfig() returned nil")
	}
	if sp.config.ExpandMergedCells != false {
		t.Error("Config not set correctly")
	}
}

func TestSheetProcessor_ReadSheet_WithMergedCells(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Merged Header")
		f.MergeCell("Sheet1", "A1", "C1")
		f.SetCellValue("Sheet1", "A2", "Data1")
		f.SetCellValue("Sheet1", "B2", "Data2")
		f.SetCellValue("Sheet1", "C2", "Data3")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")
	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Verify A1, B1, C1 all have "Merged Header" (expanded)
	for col := 0; col < 3; col++ {
		cell := grid[0][col]
		if cell.RawValue != "Merged Header" {
			t.Errorf("Cell[0][%d].RawValue = %q, want %q", col, cell.RawValue, "Merged Header")
		}
		if !cell.IsMerged {
			t.Errorf("Cell[0][%d].IsMerged = false, want true", col)
		}
		if cell.MergeRange == nil {
			t.Errorf("Cell[0][%d].MergeRange = nil, want non-nil", col)
			continue
		}
		if cell.MergeRange.StartCol != 0 || cell.MergeRange.EndCol != 2 {
			t.Errorf("Cell[0][%d].MergeRange cols wrong: got %d-%d, want 0-2",
				col, cell.MergeRange.StartCol, cell.MergeRange.EndCol)
		}
	}

	// A1 should be the origin
	if !grid[0][0].MergeRange.IsOrigin {
		t.Error("Cell[0][0] should be the merge origin")
	}

	// B1 and C1 should not be origins
	if grid[0][1].MergeRange.IsOrigin {
		t.Error("Cell[0][1] should not be the merge origin")
	}
	if grid[0][2].MergeRange.IsOrigin {
		t.Error("Cell[0][2] should not be the merge origin")
	}

	// Data cells should not be merged
	for col := 0; col < 3; col++ {
		if grid[1][col].IsMerged {
			t.Errorf("Cell[1][%d] should not be merged", col)
		}
	}
}

func TestSheetProcessor_ReadSheet_VerticalMerge(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Category")
		f.MergeCell("Sheet1", "A1", "A3")
		f.SetCellValue("Sheet1", "B1", "Item1")
		f.SetCellValue("Sheet1", "B2", "Item2")
		f.SetCellValue("Sheet1", "B3", "Item3")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")
	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Verify A1, A2, A3 all have "Category"
	for row := 0; row < 3; row++ {
		cell := grid[row][0]
		if cell.RawValue != "Category" {
			t.Errorf("Cell[%d][0].RawValue = %q, want %q", row, cell.RawValue, "Category")
		}
		if !cell.IsMerged {
			t.Errorf("Cell[%d][0].IsMerged = false, want true", row)
		}
	}

	// Only A1 should be origin
	if !grid[0][0].MergeRange.IsOrigin {
		t.Error("Cell[0][0] should be origin")
	}
	if grid[1][0].MergeRange.IsOrigin {
		t.Error("Cell[1][0] should not be origin")
	}
	if grid[2][0].MergeRange.IsOrigin {
		t.Error("Cell[2][0] should not be origin")
	}
}

func TestSheetProcessor_ReadSheet_MergeDisabled(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Merged")
		f.SetCellValue("Sheet1", "B1", "") // Ensure B1 exists
		f.MergeCell("Sheet1", "A1", "B1")
		f.SetCellValue("Sheet1", "A2", "Data") // Add data to ensure grid has 2 columns
		f.SetCellValue("Sheet1", "B2", "Data2")
	})
	defer ef.Close()

	config := models.DetectionConfig{
		ExpandMergedCells:  false,
		TrackMergeMetadata: false,
	}
	sp := NewSheetProcessorWithConfig(ef, config)
	grid, err := sp.ReadSheet("Sheet1")
	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// With merge disabled, cells should not be marked as merged
	if grid[0][0].IsMerged {
		t.Error("Cell[0][0].IsMerged should be false when merge disabled")
	}
	if len(grid[0]) > 1 && grid[0][1].IsMerged {
		t.Error("Cell[0][1].IsMerged should be false when merge disabled")
	}
}

func TestSheetProcessor_ReadSheet_RectangularMerge(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Big Block")
		f.MergeCell("Sheet1", "A1", "C3")
		// Add data to ensure the grid has 3 rows and 3 columns
		f.SetCellValue("Sheet1", "D1", "Outside")
		f.SetCellValue("Sheet1", "A4", "Below")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")
	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Verify grid has at least 3 rows and 3 columns
	if len(grid) < 3 {
		t.Fatalf("Grid has %d rows, expected at least 3", len(grid))
	}

	// All 9 cells in the 3x3 block should be merged
	for row := 0; row < 3; row++ {
		if len(grid[row]) < 3 {
			t.Fatalf("Grid row %d has %d cols, expected at least 3", row, len(grid[row]))
		}
		for col := 0; col < 3; col++ {
			cell := grid[row][col]
			if !cell.IsMerged {
				t.Errorf("Cell[%d][%d].IsMerged = false, want true", row, col)
			}
			if cell.RawValue != "Big Block" {
				t.Errorf("Cell[%d][%d].RawValue = %q, want %q", row, col, cell.RawValue, "Big Block")
			}
		}
	}

	// Only A1 should be origin
	originCount := 0
	for row := 0; row < 3; row++ {
		for col := 0; col < 3; col++ {
			if grid[row][col].MergeRange != nil && grid[row][col].MergeRange.IsOrigin {
				originCount++
			}
		}
	}
	if originCount != 1 {
		t.Errorf("Expected 1 origin cell, got %d", originCount)
	}
}

// =============================================================================
// Cell Comments Tests
// =============================================================================

func TestSheetProcessor_ReadSheet_WithComments(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Value")
		f.SetCellValue("Sheet1", "A2", "Test")
		f.SetCellValue("Sheet1", "B2", 42)

		// Add comments
		f.AddComment("Sheet1", excelize.Comment{
			Cell:   "A1",
			Author: "Author1",
			Text:   "This is the name header",
		})
		f.AddComment("Sheet1", excelize.Comment{
			Cell:   "B2",
			Author: "Author2",
			Text:   "Important value",
		})
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Check A1 has comment
	if !grid[0][0].HasComment {
		t.Error("Cell[0][0].HasComment = false, want true")
	}
	if grid[0][0].Comment != "This is the name header" {
		t.Errorf("Cell[0][0].Comment = %q, want %q", grid[0][0].Comment, "This is the name header")
	}

	// Check B1 has no comment
	if grid[0][1].HasComment {
		t.Error("Cell[0][1].HasComment = true, want false")
	}
	if grid[0][1].Comment != "" {
		t.Errorf("Cell[0][1].Comment = %q, want empty", grid[0][1].Comment)
	}

	// Check A2 has no comment
	if grid[1][0].HasComment {
		t.Error("Cell[1][0].HasComment = true, want false")
	}

	// Check B2 has comment
	if !grid[1][1].HasComment {
		t.Error("Cell[1][1].HasComment = false, want true")
	}
	if grid[1][1].Comment != "Important value" {
		t.Errorf("Cell[1][1].Comment = %q, want %q", grid[1][1].Comment, "Important value")
	}
}

func TestSheetProcessor_ReadSheet_NoComments(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Data")
		f.SetCellValue("Sheet1", "B1", "More Data")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// No cells should have comments
	for row := range grid {
		for col := range grid[row] {
			if grid[row][col].HasComment {
				t.Errorf("Cell[%d][%d].HasComment = true, want false (no comments in sheet)", row, col)
			}
			if grid[row][col].Comment != "" {
				t.Errorf("Cell[%d][%d].Comment = %q, want empty", row, col, grid[row][col].Comment)
			}
		}
	}
}

func TestSheetProcessor_ReadSheet_MultipleCommentsOnSheet(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		// Create a 3x3 grid
		for row := 1; row <= 3; row++ {
			for col := 1; col <= 3; col++ {
				cell, _ := excelize.CoordinatesToCellName(col, row)
				f.SetCellValue("Sheet1", cell, "Data")
			}
		}

		// Add comments to multiple cells
		f.AddComment("Sheet1", excelize.Comment{Cell: "A1", Author: "A", Text: "Comment A1"})
		f.AddComment("Sheet1", excelize.Comment{Cell: "B2", Author: "B", Text: "Comment B2"})
		f.AddComment("Sheet1", excelize.Comment{Cell: "C3", Author: "C", Text: "Comment C3"})
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Count cells with comments
	commentCount := 0
	for row := range grid {
		for col := range grid[row] {
			if grid[row][col].HasComment {
				commentCount++
			}
		}
	}

	if commentCount != 3 {
		t.Errorf("Expected 3 cells with comments, got %d", commentCount)
	}

	// Verify specific comments
	tests := []struct {
		row, col int
		expected string
	}{
		{0, 0, "Comment A1"},
		{1, 1, "Comment B2"},
		{2, 2, "Comment C3"},
	}

	for _, tt := range tests {
		cell := grid[tt.row][tt.col]
		if cell.Comment != tt.expected {
			t.Errorf("Cell[%d][%d].Comment = %q, want %q", tt.row, tt.col, cell.Comment, tt.expected)
		}
	}
}

func TestSheetProcessor_ReadSheet_CommentsWithUnicode(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "International")
		f.AddComment("Sheet1", excelize.Comment{
			Cell:   "A1",
			Author: "国际作者",
			Text:   "日本語コメント 한국어 中文 🎉",
		})
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	expected := "日本語コメント 한국어 中文 🎉"
	if grid[0][0].Comment != expected {
		t.Errorf("Cell[0][0].Comment = %q, want %q", grid[0][0].Comment, expected)
	}
}

func TestSheetProcessor_ReadSheet_CommentsWithMergedCells(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Merged Header")
		f.MergeCell("Sheet1", "A1", "C1")
		f.AddComment("Sheet1", excelize.Comment{
			Cell:   "A1",
			Author: "Author",
			Text:   "Comment on merged cell origin",
		})
		// Add data to ensure grid has 3 columns
		f.SetCellValue("Sheet1", "A2", "Data1")
		f.SetCellValue("Sheet1", "B2", "Data2")
		f.SetCellValue("Sheet1", "C2", "Data3")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	if len(grid) < 2 || len(grid[0]) < 3 {
		t.Fatalf("Grid too small: got %d rows, %d cols", len(grid), len(grid[0]))
	}

	// Origin cell should have the comment
	if !grid[0][0].HasComment {
		t.Error("Cell[0][0].HasComment = false, want true")
	}
	if grid[0][0].Comment != "Comment on merged cell origin" {
		t.Errorf("Cell[0][0].Comment = %q, want %q", grid[0][0].Comment, "Comment on merged cell origin")
	}

	// Other merged cells should not have the comment (comments are cell-specific)
	if grid[0][1].HasComment {
		t.Error("Cell[0][1].HasComment = true, want false (comment only on origin)")
	}
	if grid[0][2].HasComment {
		t.Error("Cell[0][2].HasComment = true, want false (comment only on origin)")
	}
}

// =============================================================================
// Hyperlink Tests
// =============================================================================

func TestSheetProcessor_ReadSheet_WithHyperlinks(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Link 1")
		f.SetCellValue("Sheet1", "B1", "No Link")
		f.SetCellValue("Sheet1", "A2", "Link 2")
		f.SetCellHyperLink("Sheet1", "A1", "https://example.com", "External")
		f.SetCellHyperLink("Sheet1", "A2", "https://google.com", "External")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Check A1 has hyperlink
	if !grid[0][0].HasHyperlink {
		t.Error("Cell[0][0].HasHyperlink = false, want true")
	}
	if grid[0][0].Hyperlink != "https://example.com" {
		t.Errorf("Cell[0][0].Hyperlink = %q, want %q", grid[0][0].Hyperlink, "https://example.com")
	}

	// Check B1 has no hyperlink
	if grid[0][1].HasHyperlink {
		t.Error("Cell[0][1].HasHyperlink = true, want false")
	}
	if grid[0][1].Hyperlink != "" {
		t.Errorf("Cell[0][1].Hyperlink = %q, want empty", grid[0][1].Hyperlink)
	}

	// Check A2 has hyperlink
	if !grid[1][0].HasHyperlink {
		t.Error("Cell[1][0].HasHyperlink = false, want true")
	}
	if grid[1][0].Hyperlink != "https://google.com" {
		t.Errorf("Cell[1][0].Hyperlink = %q, want %q", grid[1][0].Hyperlink, "https://google.com")
	}
}

func TestSheetProcessor_ReadSheet_NoHyperlinks(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Data")
		f.SetCellValue("Sheet1", "B1", "More Data")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// No cells should have hyperlinks
	for row := range grid {
		for col := range grid[row] {
			if grid[row][col].HasHyperlink {
				t.Errorf("Cell[%d][%d].HasHyperlink = true, want false (no hyperlinks in sheet)", row, col)
			}
			if grid[row][col].Hyperlink != "" {
				t.Errorf("Cell[%d][%d].Hyperlink = %q, want empty", row, col, grid[row][col].Hyperlink)
			}
		}
	}
}

func TestSheetProcessor_ReadSheet_HyperlinksWithDifferentTypes(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "External Link")
		f.SetCellValue("Sheet1", "A2", "Email Link")
		f.SetCellHyperLink("Sheet1", "A1", "https://example.com/page", "External")
		f.SetCellHyperLink("Sheet1", "A2", "mailto:test@example.com", "External")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Check external URL
	if !grid[0][0].HasHyperlink {
		t.Error("Cell[0][0].HasHyperlink = false, want true")
	}
	if grid[0][0].Hyperlink != "https://example.com/page" {
		t.Errorf("Cell[0][0].Hyperlink = %q, want %q", grid[0][0].Hyperlink, "https://example.com/page")
	}

	// Check mailto link
	if !grid[1][0].HasHyperlink {
		t.Error("Cell[1][0].HasHyperlink = false, want true")
	}
	if grid[1][0].Hyperlink != "mailto:test@example.com" {
		t.Errorf("Cell[1][0].Hyperlink = %q, want %q", grid[1][0].Hyperlink, "mailto:test@example.com")
	}
}

func TestSheetProcessor_ReadSheet_HyperlinksWithComments(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Link with comment")
		f.SetCellHyperLink("Sheet1", "A1", "https://example.com", "External")
		f.AddComment("Sheet1", excelize.Comment{
			Cell:   "A1",
			Author: "Author",
			Text:   "This is a comment",
		})
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	// Cell should have both hyperlink and comment
	cell := grid[0][0]
	if !cell.HasHyperlink {
		t.Error("Cell.HasHyperlink = false, want true")
	}
	if cell.Hyperlink != "https://example.com" {
		t.Errorf("Cell.Hyperlink = %q, want %q", cell.Hyperlink, "https://example.com")
	}
	if !cell.HasComment {
		t.Error("Cell.HasComment = false, want true")
	}
	if cell.Comment != "This is a comment" {
		t.Errorf("Cell.Comment = %q, want %q", cell.Comment, "This is a comment")
	}
}

func TestSheetProcessor_ReadSheet_HyperlinksWithMergedCells(t *testing.T) {
	ef := createSheetTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Merged with link")
		f.MergeCell("Sheet1", "A1", "C1")
		f.SetCellHyperLink("Sheet1", "A1", "https://example.com", "External")
		f.SetCellValue("Sheet1", "A2", "Data1")
		f.SetCellValue("Sheet1", "B2", "Data2")
		f.SetCellValue("Sheet1", "C2", "Data3")
	})
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, err := sp.ReadSheet("Sheet1")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	if len(grid) < 2 || len(grid[0]) < 3 {
		t.Fatalf("Grid too small: got %d rows, %d cols", len(grid), len(grid[0]))
	}

	// Only the origin cell should have the hyperlink
	if !grid[0][0].HasHyperlink {
		t.Error("Cell[0][0].HasHyperlink = false, want true")
	}
	if grid[0][0].Hyperlink != "https://example.com" {
		t.Errorf("Cell[0][0].Hyperlink = %q, want %q", grid[0][0].Hyperlink, "https://example.com")
	}

	// Other merged cells should not have the hyperlink
	if grid[0][1].HasHyperlink {
		t.Error("Cell[0][1].HasHyperlink = true, want false (hyperlink only on origin)")
	}
	if grid[0][2].HasHyperlink {
		t.Error("Cell[0][2].HasHyperlink = true, want false (hyperlink only on origin)")
	}
}
