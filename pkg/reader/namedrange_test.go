package reader

import (
	"path/filepath"
	"testing"

	"excel-lite/pkg/models"

	"github.com/xuri/excelize/v2"
)

// =============================================================================
// Test Helpers
// =============================================================================

func createNamedRangeTestFile(t *testing.T, setup func(*excelize.File)) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "test.xlsx")

	f := excelize.NewFile()
	setup(f)
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}
	return path
}

// =============================================================================
// NamedRangeReader Tests
// =============================================================================

func TestNewNamedRangeReader(t *testing.T) {
	nr := NewNamedRangeReader()
	if nr == nil {
		t.Fatal("NewNamedRangeReader() returned nil")
	}
}

func TestNewNamedRangeReaderWithConfig(t *testing.T) {
	config := models.DetectionConfig{MinRows: 5}
	nr := NewNamedRangeReaderWithConfig(config)
	if nr == nil {
		t.Fatal("NewNamedRangeReaderWithConfig() returned nil")
	}
}

func TestNamedRangeReader_GetNamedRanges(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 25)

		f.SetDefinedName(&excelize.DefinedName{
			Name:     "People",
			RefersTo: "Sheet1!$A$1:$B$2",
			Scope:    "Sheet1",
		})
		f.SetDefinedName(&excelize.DefinedName{
			Name:     "GlobalData",
			RefersTo: "Sheet1!$A$1:$B$2",
		})
	})

	nr := NewNamedRangeReader()
	ranges, err := nr.GetNamedRanges(path)

	if err != nil {
		t.Fatalf("GetNamedRanges() error = %v", err)
	}

	if len(ranges) != 2 {
		t.Errorf("GetNamedRanges() returned %d ranges, want 2", len(ranges))
	}

	// Check that both ranges are present
	foundPeople := false
	foundGlobal := false
	for _, r := range ranges {
		if r.Name == "People" {
			foundPeople = true
			if r.Scope != "Sheet1" {
				t.Errorf("People scope = %q, want %q", r.Scope, "Sheet1")
			}
		}
		if r.Name == "GlobalData" {
			foundGlobal = true
			if r.Scope != "Workbook" {
				t.Errorf("GlobalData scope = %q, want %q", r.Scope, "Workbook")
			}
		}
	}

	if !foundPeople {
		t.Error("Named range 'People' not found")
	}
	if !foundGlobal {
		t.Error("Named range 'GlobalData' not found")
	}
}

func TestNamedRangeReader_GetNamedRanges_Empty(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
	})

	nr := NewNamedRangeReader()
	ranges, err := nr.GetNamedRanges(path)

	if err != nil {
		t.Fatalf("GetNamedRanges() error = %v", err)
	}

	if len(ranges) != 0 {
		t.Errorf("GetNamedRanges() returned %d ranges, want 0", len(ranges))
	}
}

func TestNamedRangeReader_ReadRange(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		// Create data
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 25)
		f.SetCellValue("Sheet1", "A3", "Bob")
		f.SetCellValue("Sheet1", "B3", 30)

		// Create named range
		f.SetDefinedName(&excelize.DefinedName{
			Name:     "People",
			RefersTo: "Sheet1!$A$1:$B$3",
		})
	})

	nr := NewNamedRangeReader()
	table, err := nr.ReadRange(path, "People")

	if err != nil {
		t.Fatalf("ReadRange() error = %v", err)
	}

	if table == nil {
		t.Fatal("ReadRange() returned nil table")
	}

	if table.Name != "People" {
		t.Errorf("Table.Name = %q, want %q", table.Name, "People")
	}

	if len(table.Headers) != 2 {
		t.Errorf("Table has %d headers, want 2", len(table.Headers))
	}

	if len(table.Rows) != 2 {
		t.Errorf("Table has %d rows, want 2", len(table.Rows))
	}
}

func TestNamedRangeReader_ReadRange_NotFound(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
	})

	nr := NewNamedRangeReader()
	_, err := nr.ReadRange(path, "NonExistent")

	if err == nil {
		t.Error("ReadRange() expected error for non-existent range")
	}
}

func TestNamedRangeReader_ReadRange_WithQuotedSheetName(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		// Create sheet with space in name
		sheetName := "My Data"
		f.NewSheet(sheetName)
		f.SetCellValue(sheetName, "A1", "Header1")
		f.SetCellValue(sheetName, "B1", "Header2")
		f.SetCellValue(sheetName, "A2", "Value1")
		f.SetCellValue(sheetName, "B2", "Value2")

		f.SetDefinedName(&excelize.DefinedName{
			Name:     "QuotedRange",
			RefersTo: "'My Data'!$A$1:$B$2",
		})
	})

	nr := NewNamedRangeReader()
	table, err := nr.ReadRange(path, "QuotedRange")

	if err != nil {
		t.Fatalf("ReadRange() error = %v", err)
	}

	if table == nil {
		t.Fatal("ReadRange() returned nil table")
	}

	if len(table.Headers) != 2 {
		t.Errorf("Table has %d headers, want 2", len(table.Headers))
	}
}

// =============================================================================
// parseRangeReference Tests
// =============================================================================

func TestParseRangeReference(t *testing.T) {
	tests := []struct {
		name      string
		refersTo  string
		wantSheet string
		wantBound models.TableBoundary
		wantErr   bool
	}{
		{
			name:      "simple range",
			refersTo:  "Sheet1!$A$1:$B$3",
			wantSheet: "Sheet1",
			wantBound: models.TableBoundary{StartRow: 0, StartCol: 0, EndRow: 2, EndCol: 1},
		},
		{
			name:      "without dollar signs",
			refersTo:  "Sheet1!A1:B3",
			wantSheet: "Sheet1",
			wantBound: models.TableBoundary{StartRow: 0, StartCol: 0, EndRow: 2, EndCol: 1},
		},
		{
			name:      "quoted sheet name",
			refersTo:  "'My Sheet'!$A$1:$C$10",
			wantSheet: "My Sheet",
			wantBound: models.TableBoundary{StartRow: 0, StartCol: 0, EndRow: 9, EndCol: 2},
		},
		{
			name:      "large range",
			refersTo:  "Data!$A$1:$Z$100",
			wantSheet: "Data",
			wantBound: models.TableBoundary{StartRow: 0, StartCol: 0, EndRow: 99, EndCol: 25},
		},
		{
			name:     "missing sheet name",
			refersTo: "A1:B3",
			wantErr:  true,
		},
		{
			name:     "invalid format",
			refersTo: "Sheet1!A1",
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sheet, bound, err := parseRangeReference(tt.refersTo)

			if tt.wantErr {
				if err == nil {
					t.Error("parseRangeReference() expected error")
				}
				return
			}

			if err != nil {
				t.Fatalf("parseRangeReference() error = %v", err)
			}

			if sheet != tt.wantSheet {
				t.Errorf("sheet = %q, want %q", sheet, tt.wantSheet)
			}

			if bound != tt.wantBound {
				t.Errorf("boundary = %+v, want %+v", bound, tt.wantBound)
			}
		})
	}
}

// =============================================================================
// Helper Function Tests
// =============================================================================

func TestGetNamedRangeByName(t *testing.T) {
	ranges := []models.NamedRange{
		{Name: "Range1", RefersTo: "Sheet1!A1:B10"},
		{Name: "Range2", RefersTo: "Sheet1!C1:D10"},
		{Name: "Range3", RefersTo: "Sheet2!A1:B10"},
	}

	tests := []struct {
		name     string
		findName string
		want     *models.NamedRange
	}{
		{"found", "Range2", &ranges[1]},
		{"not found", "NonExistent", nil},
		{"first range", "Range1", &ranges[0]},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := GetNamedRangeByName(ranges, tt.findName)
			if tt.want == nil {
				if got != nil {
					t.Errorf("GetNamedRangeByName() = %v, want nil", got)
				}
			} else if got == nil || got.Name != tt.want.Name {
				t.Errorf("GetNamedRangeByName() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestGetNamedRangesByScope(t *testing.T) {
	ranges := []models.NamedRange{
		{Name: "Range1", Scope: "Sheet1"},
		{Name: "Range2", Scope: "Workbook"},
		{Name: "Range3", Scope: "Sheet1"},
		{Name: "Range4", Scope: "Sheet2"},
	}

	tests := []struct {
		scope    string
		expected int
	}{
		{"Sheet1", 2},
		{"Workbook", 1},
		{"Sheet2", 1},
		{"NonExistent", 0},
	}

	for _, tt := range tests {
		t.Run(tt.scope, func(t *testing.T) {
			got := GetNamedRangesByScope(ranges, tt.scope)
			if len(got) != tt.expected {
				t.Errorf("GetNamedRangesByScope(%q) returned %d ranges, want %d", tt.scope, len(got), tt.expected)
			}
		})
	}
}

func TestGetGlobalNamedRanges(t *testing.T) {
	ranges := []models.NamedRange{
		{Name: "Range1", Scope: "Sheet1"},
		{Name: "Range2", Scope: "Workbook"},
		{Name: "Range3", Scope: "Workbook"},
		{Name: "Range4", Scope: "Sheet2"},
	}

	got := GetGlobalNamedRanges(ranges)

	if len(got) != 2 {
		t.Errorf("GetGlobalNamedRanges() returned %d ranges, want 2", len(got))
	}

	for _, r := range got {
		if r.Scope != "Workbook" {
			t.Errorf("GetGlobalNamedRanges() returned range with scope %q", r.Scope)
		}
	}
}

func TestParseNamedRangeInfo(t *testing.T) {
	info, err := ParseNamedRangeInfo("Sheet1!$A$1:$C$10")

	if err != nil {
		t.Fatalf("ParseNamedRangeInfo() error = %v", err)
	}

	if info.SheetName != "Sheet1" {
		t.Errorf("SheetName = %q, want %q", info.SheetName, "Sheet1")
	}

	if info.StartRow != 0 || info.StartCol != 0 {
		t.Errorf("Start = (%d, %d), want (0, 0)", info.StartRow, info.StartCol)
	}

	if info.EndRow != 9 || info.EndCol != 2 {
		t.Errorf("End = (%d, %d), want (9, 2)", info.EndRow, info.EndCol)
	}

	if info.StartCell != "A1" {
		t.Errorf("StartCell = %q, want %q", info.StartCell, "A1")
	}

	if info.EndCell != "C10" {
		t.Errorf("EndCell = %q, want %q", info.EndCell, "C10")
	}
}

// =============================================================================
// parseCellRef Tests
// =============================================================================

func TestParseCellRef(t *testing.T) {
	tests := []struct {
		cellRef string
		wantCol int
		wantRow int
		wantErr bool
	}{
		{"A1", 0, 0, false},
		{"B2", 1, 1, false},
		{"Z26", 25, 25, false},
		{"AA1", 26, 0, false},
		{"AB10", 27, 9, false},
	}

	for _, tt := range tests {
		t.Run(tt.cellRef, func(t *testing.T) {
			col, row, err := parseCellRef(tt.cellRef)

			if tt.wantErr {
				if err == nil {
					t.Error("parseCellRef() expected error")
				}
				return
			}

			if err != nil {
				t.Fatalf("parseCellRef() error = %v", err)
			}

			if col != tt.wantCol {
				t.Errorf("col = %d, want %d", col, tt.wantCol)
			}

			if row != tt.wantRow {
				t.Errorf("row = %d, want %d", row, tt.wantRow)
			}
		})
	}
}

func TestParseColumnRow(t *testing.T) {
	tests := []struct {
		cellRef string
		wantCol int
		wantRow int
		wantErr bool
	}{
		{"A1", 0, 0, false},
		{"B2", 1, 1, false},
		{"Z26", 25, 25, false},
		{"AA1", 26, 0, false},
		{"AB10", 27, 9, false},
		{"invalid", 0, 0, true},
		{"", 0, 0, true},
	}

	for _, tt := range tests {
		t.Run(tt.cellRef, func(t *testing.T) {
			col, row, err := parseColumnRow(tt.cellRef)

			if tt.wantErr {
				if err == nil {
					t.Error("parseColumnRow() expected error")
				}
				return
			}

			if err != nil {
				t.Fatalf("parseColumnRow() error = %v", err)
			}

			if col != tt.wantCol {
				t.Errorf("col = %d, want %d", col, tt.wantCol)
			}

			if row != tt.wantRow {
				t.Errorf("row = %d, want %d", row, tt.wantRow)
			}
		})
	}
}

// =============================================================================
// Error Path Tests
// =============================================================================

func TestNamedRangeReader_GetNamedRanges_FileError(t *testing.T) {
	nr := NewNamedRangeReader()
	_, err := nr.GetNamedRanges("/nonexistent/file.xlsx")

	if err == nil {
		t.Error("GetNamedRanges() expected error for non-existent file")
	}
}

func TestNamedRangeReader_ReadRange_FileError(t *testing.T) {
	nr := NewNamedRangeReader()
	_, err := nr.ReadRange("/nonexistent/file.xlsx", "SomeRange")

	if err == nil {
		t.Error("ReadRange() expected error for non-existent file")
	}
}

func TestParseCellRef_Invalid(t *testing.T) {
	// Test invalid cell references that excelize might reject
	invalidRefs := []string{"", "123", "!@#"}

	for _, ref := range invalidRefs {
		_, _, err := parseCellRef(ref)
		if err == nil {
			t.Errorf("parseCellRef(%q) expected error", ref)
		}
	}
}

func TestParseNamedRangeInfo_Error(t *testing.T) {
	// Test invalid range references
	invalidRefs := []string{
		"NoSheet",           // Missing sheet separator
		"Sheet1!A1",         // Missing end cell
		"Sheet1!:B10",       // Missing start cell
	}

	for _, ref := range invalidRefs {
		_, err := ParseNamedRangeInfo(ref)
		if err == nil {
			t.Errorf("ParseNamedRangeInfo(%q) expected error", ref)
		}
	}
}

func TestNamedRangeReader_ReadRange_InvalidRangeReference(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
		// Create a named range with invalid reference format
		f.SetDefinedName(&excelize.DefinedName{
			Name:     "BadRange",
			RefersTo: "InvalidRef",
		})
	})

	nr := NewNamedRangeReader()
	_, err := nr.ReadRange(path, "BadRange")

	if err == nil {
		t.Error("ReadRange() expected error for invalid range reference")
	}
}

// =============================================================================
// Integration Tests
// =============================================================================

func TestNamedRangeReader_ReadRange_LargeRange(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		// Create a 10x10 table
		headers := []string{"Col1", "Col2", "Col3", "Col4", "Col5"}
		for i, h := range headers {
			cell, _ := excelize.CoordinatesToCellName(i+1, 1)
			f.SetCellValue("Sheet1", cell, h)
		}

		for row := 2; row <= 11; row++ {
			for col := 1; col <= 5; col++ {
				cell, _ := excelize.CoordinatesToCellName(col, row)
				f.SetCellValue("Sheet1", cell, (row-1)*col)
			}
		}

		f.SetDefinedName(&excelize.DefinedName{
			Name:     "BigTable",
			RefersTo: "Sheet1!$A$1:$E$11",
		})
	})

	nr := NewNamedRangeReader()
	table, err := nr.ReadRange(path, "BigTable")

	if err != nil {
		t.Fatalf("ReadRange() error = %v", err)
	}

	if len(table.Headers) != 5 {
		t.Errorf("Table has %d headers, want 5", len(table.Headers))
	}

	if len(table.Rows) != 10 {
		t.Errorf("Table has %d rows, want 10", len(table.Rows))
	}
}

func TestNamedRangeReader_ReadRange_PartialRange(t *testing.T) {
	path := createNamedRangeTestFile(t, func(f *excelize.File) {
		// Create data in columns B-D (not starting at A)
		f.SetCellValue("Sheet1", "B2", "Name")
		f.SetCellValue("Sheet1", "C2", "Age")
		f.SetCellValue("Sheet1", "D2", "City")
		f.SetCellValue("Sheet1", "B3", "Alice")
		f.SetCellValue("Sheet1", "C3", 25)
		f.SetCellValue("Sheet1", "D3", "NYC")

		f.SetDefinedName(&excelize.DefinedName{
			Name:     "PartialRange",
			RefersTo: "Sheet1!$B$2:$D$3",
		})
	})

	nr := NewNamedRangeReader()
	table, err := nr.ReadRange(path, "PartialRange")

	if err != nil {
		t.Fatalf("ReadRange() error = %v", err)
	}

	if len(table.Headers) != 3 {
		t.Errorf("Table has %d headers, want 3", len(table.Headers))
	}

	if len(table.Rows) != 1 {
		t.Errorf("Table has %d rows, want 1", len(table.Rows))
	}

	// Check header names
	expectedHeaders := []string{"Name", "Age", "City"}
	for i, h := range expectedHeaders {
		if i < len(table.Headers) && table.Headers[i] != h {
			t.Errorf("Header[%d] = %q, want %q", i, table.Headers[i], h)
		}
	}
}
