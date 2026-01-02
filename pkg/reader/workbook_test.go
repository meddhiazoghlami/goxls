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

func createWorkbookTestFile(t *testing.T, setup func(*excelize.File)) string {
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
// Constructor Tests
// =============================================================================

func TestNewWorkbookReader(t *testing.T) {
	wr := NewWorkbookReader()

	if wr == nil {
		t.Fatal("NewWorkbookReader() returned nil")
	}

	if wr.analyzer == nil {
		t.Error("analyzer is nil")
	}

	if wr.headerDetector == nil {
		t.Error("headerDetector is nil")
	}

	if wr.rowParser == nil {
		t.Error("rowParser is nil")
	}
}

func TestNewWorkbookReaderWithConfig(t *testing.T) {
	config := models.DetectionConfig{
		MinColumns: 5,
		MinRows:    10,
	}

	wr := NewWorkbookReaderWithConfig(config)

	if wr == nil {
		t.Fatal("NewWorkbookReaderWithConfig() returned nil")
	}

	if wr.config.MinColumns != 5 {
		t.Errorf("config.MinColumns = %d, want 5", wr.config.MinColumns)
	}

	if wr.config.MinRows != 10 {
		t.Errorf("config.MinRows = %d, want 10", wr.config.MinRows)
	}
}

// =============================================================================
// ReadFile Tests
// =============================================================================

func TestWorkbookReader_ReadFile_Simple(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 30)
		f.SetCellValue("Sheet1", "A3", "Bob")
		f.SetCellValue("Sheet1", "B3", 25)
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	if wb == nil {
		t.Fatal("ReadFile() returned nil workbook")
	}

	if wb.FilePath != path {
		t.Errorf("wb.FilePath = %q, want %q", wb.FilePath, path)
	}

	if len(wb.Sheets) != 1 {
		t.Fatalf("len(wb.Sheets) = %d, want 1", len(wb.Sheets))
	}

	if len(wb.Sheets[0].Tables) == 0 {
		t.Fatal("No tables detected")
	}
}

func TestWorkbookReader_ReadFile_MultipleSheets(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Sheet 1
		f.SetCellValue("Sheet1", "A1", "ID")
		f.SetCellValue("Sheet1", "B1", "Value")
		f.SetCellValue("Sheet1", "A2", "1")
		f.SetCellValue("Sheet1", "B2", "100")

		// Sheet 2
		f.NewSheet("Sheet2")
		f.SetCellValue("Sheet2", "A1", "Name")
		f.SetCellValue("Sheet2", "B1", "Score")
		f.SetCellValue("Sheet2", "A2", "Test")
		f.SetCellValue("Sheet2", "B2", "95")

		// Sheet 3
		f.NewSheet("Sheet3")
		f.SetCellValue("Sheet3", "A1", "Data")
		f.SetCellValue("Sheet3", "B1", "More")
		f.SetCellValue("Sheet3", "A2", "X")
		f.SetCellValue("Sheet3", "B2", "Y")
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	if len(wb.Sheets) != 3 {
		t.Errorf("len(wb.Sheets) = %d, want 3", len(wb.Sheets))
	}

	// Check sheet names
	expectedNames := []string{"Sheet1", "Sheet2", "Sheet3"}
	for i, name := range expectedNames {
		if wb.Sheets[i].Name != name {
			t.Errorf("wb.Sheets[%d].Name = %q, want %q", i, wb.Sheets[i].Name, name)
		}
	}
}

func TestWorkbookReader_ReadFile_EmptySheet(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Empty sheet
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	if len(wb.Sheets) != 1 {
		t.Fatalf("len(wb.Sheets) = %d, want 1", len(wb.Sheets))
	}

	if len(wb.Sheets[0].Tables) != 0 {
		t.Errorf("Empty sheet should have 0 tables, got %d", len(wb.Sheets[0].Tables))
	}
}

func TestWorkbookReader_ReadFile_FileNotFound(t *testing.T) {
	wr := NewWorkbookReader()
	_, err := wr.ReadFile("/nonexistent/path/file.xlsx")

	if err == nil {
		t.Error("ReadFile() expected error for nonexistent file")
	}
}

func TestWorkbookReader_ReadFile_MultipleTables(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Table 1: rows 1-4
		f.SetCellValue("Sheet1", "A1", "Table1_Col1")
		f.SetCellValue("Sheet1", "B1", "Table1_Col2")
		f.SetCellValue("Sheet1", "A2", "Data1")
		f.SetCellValue("Sheet1", "B2", "Data2")
		f.SetCellValue("Sheet1", "A3", "Data3")
		f.SetCellValue("Sheet1", "B3", "Data4")

		// Gap (rows 5-7 empty)

		// Table 2: rows 8-10
		f.SetCellValue("Sheet1", "A8", "Table2_Col1")
		f.SetCellValue("Sheet1", "B8", "Table2_Col2")
		f.SetCellValue("Sheet1", "C8", "Table2_Col3")
		f.SetCellValue("Sheet1", "A9", "A")
		f.SetCellValue("Sheet1", "B9", "B")
		f.SetCellValue("Sheet1", "C9", "C")
		f.SetCellValue("Sheet1", "A10", "D")
		f.SetCellValue("Sheet1", "B10", "E")
		f.SetCellValue("Sheet1", "C10", "F")
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	// Should detect 2 tables
	if len(wb.Sheets[0].Tables) < 2 {
		t.Errorf("Expected at least 2 tables, got %d", len(wb.Sheets[0].Tables))
	}
}

func TestWorkbookReader_ReadFile_OffsetTable(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Metadata in top-left
		f.SetCellValue("Sheet1", "A1", "Report Title")

		// Table starting at B3
		f.SetCellValue("Sheet1", "B3", "Name")
		f.SetCellValue("Sheet1", "C3", "Value")
		f.SetCellValue("Sheet1", "B4", "Item1")
		f.SetCellValue("Sheet1", "C4", "100")
		f.SetCellValue("Sheet1", "B5", "Item2")
		f.SetCellValue("Sheet1", "C5", "200")
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	// Should detect table
	if len(wb.Sheets[0].Tables) == 0 {
		t.Fatal("No tables detected for offset table")
	}
}

// =============================================================================
// ReadSheet Tests
// =============================================================================

func TestWorkbookReader_ReadSheet(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 30)

		f.NewSheet("Sheet2")
		f.SetCellValue("Sheet2", "A1", "Other")
		f.SetCellValue("Sheet2", "B1", "Data")
	})

	wr := NewWorkbookReader()
	sheet, err := wr.ReadSheet(path, "Sheet2")

	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	if sheet == nil {
		t.Fatal("ReadSheet() returned nil")
	}

	if sheet.Name != "Sheet2" {
		t.Errorf("sheet.Name = %q, want Sheet2", sheet.Name)
	}
}

func TestWorkbookReader_ReadSheet_NotFound(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Data")
	})

	wr := NewWorkbookReader()
	_, err := wr.ReadSheet(path, "NonexistentSheet")

	if err == nil {
		t.Error("ReadSheet() expected error for nonexistent sheet")
	}
}

// =============================================================================
// Helper Functions Tests
// =============================================================================

func TestGetTableByName(t *testing.T) {
	wb := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Sheet1",
				Tables: []models.Table{
					{Name: "Sheet1_Table1", Headers: []string{"A", "B"}},
					{Name: "Sheet1_Table2", Headers: []string{"C", "D"}},
				},
			},
			{
				Name: "Sheet2",
				Tables: []models.Table{
					{Name: "Sheet2_Table1", Headers: []string{"E", "F"}},
				},
			},
		},
	}

	tests := []struct {
		name      string
		tableName string
		found     bool
	}{
		{"first table", "Sheet1_Table1", true},
		{"second table", "Sheet1_Table2", true},
		{"table on second sheet", "Sheet2_Table1", true},
		{"nonexistent", "NotATable", false},
		{"empty name", "", false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			table := GetTableByName(wb, tt.tableName)
			if tt.found && table == nil {
				t.Errorf("GetTableByName(%q) returned nil, expected table", tt.tableName)
			}
			if !tt.found && table != nil {
				t.Errorf("GetTableByName(%q) returned table, expected nil", tt.tableName)
			}
			if tt.found && table != nil && table.Name != tt.tableName {
				t.Errorf("GetTableByName(%q) returned table with name %q", tt.tableName, table.Name)
			}
		})
	}
}

func TestGetTableByName_NilWorkbook(t *testing.T) {
	table := GetTableByName(nil, "AnyTable")
	if table != nil {
		t.Error("GetTableByName(nil, ...) should return nil")
	}
}

func TestGetAllTables(t *testing.T) {
	wb := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Tables: []models.Table{
					{Name: "T1"},
					{Name: "T2"},
				},
			},
			{
				Tables: []models.Table{
					{Name: "T3"},
				},
			},
			{
				Tables: []models.Table{}, // Empty
			},
		},
	}

	tables := GetAllTables(wb)

	if len(tables) != 3 {
		t.Errorf("GetAllTables() returned %d tables, want 3", len(tables))
	}

	names := make(map[string]bool)
	for _, table := range tables {
		names[table.Name] = true
	}

	for _, expected := range []string{"T1", "T2", "T3"} {
		if !names[expected] {
			t.Errorf("GetAllTables() missing table %q", expected)
		}
	}
}

func TestGetAllTables_Empty(t *testing.T) {
	wb := &models.Workbook{
		Sheets: []models.Sheet{},
	}

	tables := GetAllTables(wb)

	if len(tables) != 0 {
		t.Errorf("GetAllTables() on empty workbook returned %d tables, want 0", len(tables))
	}
}

func TestGetAllTables_NilWorkbook(t *testing.T) {
	tables := GetAllTables(nil)
	if tables != nil {
		t.Error("GetAllTables(nil) should return nil")
	}
}

// =============================================================================
// Integration Tests
// =============================================================================

func TestWorkbookReader_Integration_FullWorkflow(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Create a realistic workbook
		f.SetSheetName("Sheet1", "Employees")

		// Employee table
		f.SetCellValue("Employees", "A1", "ID")
		f.SetCellValue("Employees", "B1", "Name")
		f.SetCellValue("Employees", "C1", "Department")
		f.SetCellValue("Employees", "D1", "Salary")

		f.SetCellValue("Employees", "A2", 1)
		f.SetCellValue("Employees", "B2", "Alice Johnson")
		f.SetCellValue("Employees", "C2", "Engineering")
		f.SetCellValue("Employees", "D2", 75000)

		f.SetCellValue("Employees", "A3", 2)
		f.SetCellValue("Employees", "B3", "Bob Smith")
		f.SetCellValue("Employees", "C3", "Marketing")
		f.SetCellValue("Employees", "D3", 65000)

		f.SetCellValue("Employees", "A4", 3)
		f.SetCellValue("Employees", "B4", "Carol White")
		f.SetCellValue("Employees", "C4", "Engineering")
		f.SetCellValue("Employees", "D4", 80000)
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	// Verify workbook structure
	if len(wb.Sheets) != 1 {
		t.Fatalf("Expected 1 sheet, got %d", len(wb.Sheets))
	}

	sheet := wb.Sheets[0]
	if sheet.Name != "Employees" {
		t.Errorf("Sheet name = %q, want Employees", sheet.Name)
	}

	if len(sheet.Tables) == 0 {
		t.Fatal("No tables detected")
	}

	table := sheet.Tables[0]

	// Verify headers
	expectedHeaders := []string{"ID", "Name", "Department", "Salary"}
	if len(table.Headers) < len(expectedHeaders) {
		t.Errorf("Expected at least %d headers, got %d", len(expectedHeaders), len(table.Headers))
	}

	// Verify data rows
	if len(table.Rows) < 3 {
		t.Errorf("Expected at least 3 data rows, got %d", len(table.Rows))
	}

	// Verify we can access data
	if len(table.Rows) > 0 {
		firstRow := table.Rows[0]
		if nameCell, ok := firstRow.Get("Name"); ok {
			if nameCell.AsString() != "Alice Johnson" {
				t.Errorf("First row Name = %q, want Alice Johnson", nameCell.AsString())
			}
		}
	}
}

func TestWorkbookReader_Integration_ComplexLayout(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Title and metadata
		f.SetCellValue("Sheet1", "A1", "Monthly Report")
		f.SetCellValue("Sheet1", "A2", "Generated: 2024-01-15")

		// Data table starting at row 4
		f.SetCellValue("Sheet1", "A4", "Region")
		f.SetCellValue("Sheet1", "B4", "Q1 Sales")
		f.SetCellValue("Sheet1", "C4", "Q2 Sales")
		f.SetCellValue("Sheet1", "D4", "Total")

		f.SetCellValue("Sheet1", "A5", "North")
		f.SetCellValue("Sheet1", "B5", 10000)
		f.SetCellValue("Sheet1", "C5", 12000)
		f.SetCellValue("Sheet1", "D5", 22000)

		f.SetCellValue("Sheet1", "A6", "South")
		f.SetCellValue("Sheet1", "B6", 8000)
		f.SetCellValue("Sheet1", "C6", 9500)
		f.SetCellValue("Sheet1", "D6", 17500)

		f.SetCellValue("Sheet1", "A7", "East")
		f.SetCellValue("Sheet1", "B7", 15000)
		f.SetCellValue("Sheet1", "C7", 14000)
		f.SetCellValue("Sheet1", "D7", 29000)
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	// Should detect at least one table
	if len(wb.Sheets[0].Tables) == 0 {
		t.Fatal("No tables detected in complex layout")
	}

	// Find the main data table (the one with Region header)
	var mainTable *models.Table
	for i := range wb.Sheets[0].Tables {
		table := &wb.Sheets[0].Tables[i]
		for _, h := range table.Headers {
			if h == "Region" {
				mainTable = table
				break
			}
		}
	}

	if mainTable == nil {
		t.Fatal("Main data table with 'Region' header not found")
	}

	if len(mainTable.Rows) < 3 {
		t.Errorf("Main table should have at least 3 rows, got %d", len(mainTable.Rows))
	}
}

func TestWorkbookReader_Integration_UnicodeContent(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "名前")
		f.SetCellValue("Sheet1", "B1", "都市")
		f.SetCellValue("Sheet1", "A2", "田中太郎")
		f.SetCellValue("Sheet1", "B2", "東京")
		f.SetCellValue("Sheet1", "A3", "鈴木花子")
		f.SetCellValue("Sheet1", "B3", "大阪")
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	if len(wb.Sheets[0].Tables) == 0 {
		t.Fatal("No tables detected")
	}

	table := wb.Sheets[0].Tables[0]

	// Check headers contain unicode
	foundUnicodeHeader := false
	for _, h := range table.Headers {
		if h == "名前" || h == "都市" {
			foundUnicodeHeader = true
			break
		}
	}

	if !foundUnicodeHeader {
		t.Error("Unicode headers not properly detected")
	}
}

func TestWorkbookReader_Integration_LargeFile(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping large file test in short mode")
	}

	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Create 500 rows x 20 columns
		for col := 1; col <= 20; col++ {
			cell, _ := excelize.CoordinatesToCellName(col, 1)
			f.SetCellValue("Sheet1", cell, "Header"+string(rune('A'-1+col)))
		}

		for row := 2; row <= 500; row++ {
			for col := 1; col <= 20; col++ {
				cell, _ := excelize.CoordinatesToCellName(col, row)
				f.SetCellValue("Sheet1", cell, row*col)
			}
		}
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFile(path)

	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	if len(wb.Sheets[0].Tables) == 0 {
		t.Fatal("No tables detected in large file")
	}

	table := wb.Sheets[0].Tables[0]
	if len(table.Rows) < 400 {
		t.Errorf("Expected at least 400 rows, got %d", len(table.Rows))
	}
}

func TestWorkbookReader_ReadSheet_FileNotFound(t *testing.T) {
	wr := NewWorkbookReader()
	_, err := wr.ReadSheet("/nonexistent/path/file.xlsx", "Sheet1")

	if err == nil {
		t.Error("ReadSheet() expected error for nonexistent file")
	}
}

func TestWorkbookReader_ReadSheet_SheetIndex(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Data1")
		f.NewSheet("Sheet2")
		f.SetCellValue("Sheet2", "A1", "Data2")
		f.NewSheet("Sheet3")
		f.SetCellValue("Sheet3", "A1", "Data3")
	})

	wr := NewWorkbookReader()

	// Read second sheet
	sheet, err := wr.ReadSheet(path, "Sheet2")
	if err != nil {
		t.Fatalf("ReadSheet() error = %v", err)
	}

	if sheet.Index != 1 {
		t.Errorf("sheet.Index = %d, want 1", sheet.Index)
	}
}

func TestWorkbookReader_ReadFile_InvalidFile(t *testing.T) {
	wr := NewWorkbookReader()
	_, err := wr.ReadFile("/nonexistent/totally/fake/path.xlsx")
	if err == nil {
		t.Error("Expected error for invalid file")
	}
}

// =============================================================================
// ReadFileParallel Tests
// =============================================================================

func TestWorkbookReader_ReadFileParallel_MultipleSheets(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Sheet 1
		f.SetCellValue("Sheet1", "A1", "ID")
		f.SetCellValue("Sheet1", "B1", "Value")
		f.SetCellValue("Sheet1", "A2", "1")
		f.SetCellValue("Sheet1", "B2", "100")
		f.SetCellValue("Sheet1", "A3", "2")
		f.SetCellValue("Sheet1", "B3", "200")

		// Sheet 2
		f.NewSheet("Sheet2")
		f.SetCellValue("Sheet2", "A1", "Name")
		f.SetCellValue("Sheet2", "B1", "Score")
		f.SetCellValue("Sheet2", "A2", "Alice")
		f.SetCellValue("Sheet2", "B2", "95")
		f.SetCellValue("Sheet2", "A3", "Bob")
		f.SetCellValue("Sheet2", "B3", "87")

		// Sheet 3
		f.NewSheet("Sheet3")
		f.SetCellValue("Sheet3", "A1", "Product")
		f.SetCellValue("Sheet3", "B1", "Price")
		f.SetCellValue("Sheet3", "A2", "Widget")
		f.SetCellValue("Sheet3", "B2", "9.99")
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFileParallel(path)

	if err != nil {
		t.Fatalf("ReadFileParallel() error = %v", err)
	}

	if wb == nil {
		t.Fatal("ReadFileParallel() returned nil workbook")
	}

	if len(wb.Sheets) != 3 {
		t.Errorf("len(wb.Sheets) = %d, want 3", len(wb.Sheets))
	}

	// Verify sheet order is maintained
	expectedNames := []string{"Sheet1", "Sheet2", "Sheet3"}
	for i, name := range expectedNames {
		if wb.Sheets[i].Name != name {
			t.Errorf("wb.Sheets[%d].Name = %q, want %q", i, wb.Sheets[i].Name, name)
		}
	}

	// Verify each sheet has tables
	for i, sheet := range wb.Sheets {
		if len(sheet.Tables) == 0 {
			t.Errorf("Sheet %d (%s) has no tables", i, sheet.Name)
		}
	}
}

func TestWorkbookReader_ReadFileParallel_SingleSheet(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 30)
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFileParallel(path)

	if err != nil {
		t.Fatalf("ReadFileParallel() error = %v", err)
	}

	if len(wb.Sheets) != 1 {
		t.Errorf("len(wb.Sheets) = %d, want 1", len(wb.Sheets))
	}

	if len(wb.Sheets[0].Tables) == 0 {
		t.Error("No tables detected in single sheet")
	}
}

func TestWorkbookReader_ReadFileParallel_EmptySheet(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Empty sheet
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFileParallel(path)

	if err != nil {
		t.Fatalf("ReadFileParallel() error = %v", err)
	}

	if len(wb.Sheets) != 1 {
		t.Errorf("len(wb.Sheets) = %d, want 1", len(wb.Sheets))
	}

	if len(wb.Sheets[0].Tables) != 0 {
		t.Errorf("Empty sheet should have 0 tables, got %d", len(wb.Sheets[0].Tables))
	}
}

func TestWorkbookReader_ReadFileParallel_FileNotFound(t *testing.T) {
	wr := NewWorkbookReader()
	_, err := wr.ReadFileParallel("/nonexistent/path/file.xlsx")

	if err == nil {
		t.Error("ReadFileParallel() expected error for nonexistent file")
	}
}

func TestWorkbookReader_ReadFileParallel_MatchesSequential(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Sheet 1
		f.SetCellValue("Sheet1", "A1", "ID")
		f.SetCellValue("Sheet1", "B1", "Name")
		f.SetCellValue("Sheet1", "A2", "1")
		f.SetCellValue("Sheet1", "B2", "Alice")
		f.SetCellValue("Sheet1", "A3", "2")
		f.SetCellValue("Sheet1", "B3", "Bob")

		// Sheet 2
		f.NewSheet("Data")
		f.SetCellValue("Data", "A1", "X")
		f.SetCellValue("Data", "B1", "Y")
		f.SetCellValue("Data", "A2", "10")
		f.SetCellValue("Data", "B2", "20")
	})

	wr := NewWorkbookReader()

	// Read sequentially
	wbSeq, err := wr.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	// Read in parallel
	wbPar, err := wr.ReadFileParallel(path)
	if err != nil {
		t.Fatalf("ReadFileParallel() error = %v", err)
	}

	// Compare results
	if len(wbSeq.Sheets) != len(wbPar.Sheets) {
		t.Fatalf("Sheet count mismatch: sequential=%d, parallel=%d",
			len(wbSeq.Sheets), len(wbPar.Sheets))
	}

	for i := range wbSeq.Sheets {
		seqSheet := wbSeq.Sheets[i]
		parSheet := wbPar.Sheets[i]

		if seqSheet.Name != parSheet.Name {
			t.Errorf("Sheet %d name mismatch: sequential=%q, parallel=%q",
				i, seqSheet.Name, parSheet.Name)
		}

		if seqSheet.Index != parSheet.Index {
			t.Errorf("Sheet %d index mismatch: sequential=%d, parallel=%d",
				i, seqSheet.Index, parSheet.Index)
		}

		if len(seqSheet.Tables) != len(parSheet.Tables) {
			t.Errorf("Sheet %d table count mismatch: sequential=%d, parallel=%d",
				i, len(seqSheet.Tables), len(parSheet.Tables))
			continue
		}

		for j := range seqSheet.Tables {
			seqTable := seqSheet.Tables[j]
			parTable := parSheet.Tables[j]

			if seqTable.Name != parTable.Name {
				t.Errorf("Sheet %d, Table %d name mismatch: sequential=%q, parallel=%q",
					i, j, seqTable.Name, parTable.Name)
			}

			if len(seqTable.Headers) != len(parTable.Headers) {
				t.Errorf("Sheet %d, Table %d header count mismatch: sequential=%d, parallel=%d",
					i, j, len(seqTable.Headers), len(parTable.Headers))
			}

			if len(seqTable.Rows) != len(parTable.Rows) {
				t.Errorf("Sheet %d, Table %d row count mismatch: sequential=%d, parallel=%d",
					i, j, len(seqTable.Rows), len(parTable.Rows))
			}
		}
	}
}

func TestWorkbookReader_ReadFileParallel_ManySheets(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping many sheets test in short mode")
	}

	path := createWorkbookTestFile(t, func(f *excelize.File) {
		// Create 5 sheets with data
		sheetNames := []string{"Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"}

		for i, sheetName := range sheetNames {
			if i > 0 {
				f.NewSheet(sheetName)
			}

			// Add data to each sheet
			f.SetCellValue(sheetName, "A1", "Header"+sheetName)
			f.SetCellValue(sheetName, "B1", "Value")
			f.SetCellValue(sheetName, "A2", "Data")
			f.SetCellValue(sheetName, "B2", (i+1)*100)
		}
	})

	wr := NewWorkbookReader()
	wb, err := wr.ReadFileParallel(path)

	if err != nil {
		t.Fatalf("ReadFileParallel() error = %v", err)
	}

	// Should have processed all 5 sheets
	if len(wb.Sheets) != 5 {
		t.Errorf("Expected 5 sheets, got %d", len(wb.Sheets))
	}
}

func TestWorkbookReader_ReadFileParallel_WithCustomConfig(t *testing.T) {
	path := createWorkbookTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Col1")
		f.SetCellValue("Sheet1", "B1", "Col2")
		f.SetCellValue("Sheet1", "A2", "Data1")
		f.SetCellValue("Sheet1", "B2", "Data2")

		f.NewSheet("Sheet2")
		f.SetCellValue("Sheet2", "A1", "X")
		f.SetCellValue("Sheet2", "B1", "Y")
		f.SetCellValue("Sheet2", "A2", "1")
		f.SetCellValue("Sheet2", "B2", "2")
	})

	config := models.DetectionConfig{
		MinColumns:         2,
		MinRows:            2,
		MaxEmptyRows:       2,
		HeaderDensity:      0.5,
		ColumnConsistency:  0.7,
		ExpandMergedCells:  true,
		TrackMergeMetadata: true,
	}

	wr := NewWorkbookReaderWithConfig(config)
	wb, err := wr.ReadFileParallel(path)

	if err != nil {
		t.Fatalf("ReadFileParallel() error = %v", err)
	}

	if len(wb.Sheets) != 2 {
		t.Errorf("len(wb.Sheets) = %d, want 2", len(wb.Sheets))
	}
}
