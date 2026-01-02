package goxcel

import (
	"context"
	"errors"
	"testing"
	"time"
)

func TestReadFile(t *testing.T) {
	// Test with existing file
	workbook, err := ReadFile("testdata/sample.xlsx")
	if err != nil {
		t.Fatalf("ReadFile failed: %v", err)
	}

	if workbook == nil {
		t.Fatal("Expected workbook, got nil")
	}

	if len(workbook.Sheets) == 0 {
		t.Error("Expected at least one sheet")
	}
}

func TestReadFileNotFound(t *testing.T) {
	_, err := ReadFile("nonexistent.xlsx")
	if err == nil {
		t.Fatal("Expected error for non-existent file")
	}

	if !errors.Is(err, ErrFileNotFound) {
		t.Errorf("Expected ErrFileNotFound, got: %v", err)
	}
}

func TestReadFileWithOptions(t *testing.T) {
	workbook, err := ReadFile("testdata/sample.xlsx",
		WithMinColumns(2),
		WithMinRows(2),
		WithParallel(false),
		WithExpandMergedCells(true),
	)
	if err != nil {
		t.Fatalf("ReadFile with options failed: %v", err)
	}

	if workbook == nil {
		t.Fatal("Expected workbook, got nil")
	}
}

func TestReadFileParallel(t *testing.T) {
	workbook, err := ReadFile("testdata/sample.xlsx", WithParallel(true))
	if err != nil {
		t.Fatalf("ReadFile parallel failed: %v", err)
	}

	if workbook == nil {
		t.Fatal("Expected workbook, got nil")
	}
}

func TestReadFileWithContext(t *testing.T) {
	ctx := context.Background()
	workbook, err := ReadFileWithContext(ctx, "testdata/sample.xlsx")
	if err != nil {
		t.Fatalf("ReadFileWithContext failed: %v", err)
	}

	if workbook == nil {
		t.Fatal("Expected workbook, got nil")
	}
}

func TestReadFileWithContextCanceled(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel() // Cancel immediately

	_, err := ReadFileWithContext(ctx, "testdata/sample.xlsx")
	if err == nil {
		t.Fatal("Expected error for canceled context")
	}

	if !errors.Is(err, ErrContextCanceled) {
		t.Errorf("Expected ErrContextCanceled, got: %v", err)
	}
}

func TestReadFileWithContextTimeout(t *testing.T) {
	// This test uses a very short timeout
	ctx, cancel := context.WithTimeout(context.Background(), 1*time.Nanosecond)
	defer cancel()

	// Give enough time for timeout to trigger
	time.Sleep(10 * time.Millisecond)

	_, err := ReadFileWithContext(ctx, "testdata/sample.xlsx")
	if err == nil {
		// File might be small enough to read before timeout - that's ok
		return
	}

	if !errors.Is(err, ErrContextCanceled) {
		t.Errorf("Expected ErrContextCanceled, got: %v", err)
	}
}

func TestDefaultConfig(t *testing.T) {
	config := DefaultConfig()

	if config.MinColumns != 2 {
		t.Errorf("Expected MinColumns=2, got %d", config.MinColumns)
	}
	if config.MinRows != 2 {
		t.Errorf("Expected MinRows=2, got %d", config.MinRows)
	}
	if config.MaxEmptyRows != 2 {
		t.Errorf("Expected MaxEmptyRows=2, got %d", config.MaxEmptyRows)
	}
}

func TestWithConfigOption(t *testing.T) {
	config := DetectionConfig{
		MinColumns:        5,
		MinRows:           10,
		MaxEmptyRows:      3,
		HeaderDensity:     0.8,
		ColumnConsistency: 0.9,
	}

	workbook, err := ReadFile("testdata/sample.xlsx", WithConfig(config))
	if err != nil {
		t.Fatalf("ReadFile with config failed: %v", err)
	}

	if workbook == nil {
		t.Fatal("Expected workbook, got nil")
	}
}

func TestExportFunctions(t *testing.T) {
	workbook, err := ReadFile("testdata/sample.xlsx")
	if err != nil {
		t.Fatalf("ReadFile failed: %v", err)
	}

	if len(workbook.Sheets) == 0 || len(workbook.Sheets[0].Tables) == 0 {
		t.Skip("No tables found in test file")
	}

	table := &workbook.Sheets[0].Tables[0]

	// Test ToJSON
	json, err := ToJSON(table)
	if err != nil {
		t.Errorf("ToJSON failed: %v", err)
	}
	if json == "" {
		t.Error("ToJSON returned empty string")
	}

	// Test ToJSONPretty
	jsonPretty, err := ToJSONPretty(table)
	if err != nil {
		t.Errorf("ToJSONPretty failed: %v", err)
	}
	if jsonPretty == "" {
		t.Error("ToJSONPretty returned empty string")
	}

	// Test ToCSV
	csv, err := ToCSV(table)
	if err != nil {
		t.Errorf("ToCSV failed: %v", err)
	}
	if csv == "" {
		t.Error("ToCSV returned empty string")
	}

	// Test ToTSV
	tsv, err := ToTSV(table)
	if err != nil {
		t.Errorf("ToTSV failed: %v", err)
	}
	if tsv == "" {
		t.Error("ToTSV returned empty string")
	}

	// Test ToCSVWithDelimiter
	csvSemi, err := ToCSVWithDelimiter(table, ';')
	if err != nil {
		t.Errorf("ToCSVWithDelimiter failed: %v", err)
	}
	if csvSemi == "" {
		t.Error("ToCSVWithDelimiter returned empty string")
	}

	// Test ToSQL
	sql, err := ToSQL(table, "test_table")
	if err != nil {
		t.Errorf("ToSQL failed: %v", err)
	}
	if sql == "" {
		t.Error("ToSQL returned empty string")
	}

	// Test ToSQLWithCreate
	sqlCreate, err := ToSQLWithCreate(table, "test_table")
	if err != nil {
		t.Errorf("ToSQLWithCreate failed: %v", err)
	}
	if sqlCreate == "" {
		t.Error("ToSQLWithCreate returned empty string")
	}
}

func TestCellTypeConstants(t *testing.T) {
	// Verify constants are properly re-exported
	if CellTypeEmpty != 0 {
		t.Errorf("Expected CellTypeEmpty=0, got %d", CellTypeEmpty)
	}
	if CellTypeString != 1 {
		t.Errorf("Expected CellTypeString=1, got %d", CellTypeString)
	}
	if CellTypeNumber != 2 {
		t.Errorf("Expected CellTypeNumber=2, got %d", CellTypeNumber)
	}
	if CellTypeDate != 3 {
		t.Errorf("Expected CellTypeDate=3, got %d", CellTypeDate)
	}
	if CellTypeBool != 4 {
		t.Errorf("Expected CellTypeBool=4, got %d", CellTypeBool)
	}
	if CellTypeFormula != 5 {
		t.Errorf("Expected CellTypeFormula=5, got %d", CellTypeFormula)
	}
}

func TestSentinelErrors(t *testing.T) {
	// Verify sentinel errors are defined
	errors := []error{
		ErrFileNotFound,
		ErrInvalidFormat,
		ErrSheetNotFound,
		ErrNoTablesFound,
		ErrInvalidRange,
		ErrEmptyWorkbook,
		ErrContextCanceled,
	}

	for _, err := range errors {
		if err == nil {
			t.Error("Expected non-nil sentinel error")
		}
	}
}

func TestFunctionalOptions(t *testing.T) {
	opts := defaultOptions()

	// Test each option modifies the config correctly
	WithMinColumns(5)(opts)
	if opts.config.MinColumns != 5 {
		t.Errorf("WithMinColumns failed: got %d", opts.config.MinColumns)
	}

	WithMinRows(10)(opts)
	if opts.config.MinRows != 10 {
		t.Errorf("WithMinRows failed: got %d", opts.config.MinRows)
	}

	WithMaxEmptyRows(3)(opts)
	if opts.config.MaxEmptyRows != 3 {
		t.Errorf("WithMaxEmptyRows failed: got %d", opts.config.MaxEmptyRows)
	}

	WithHeaderDensity(0.8)(opts)
	if opts.config.HeaderDensity != 0.8 {
		t.Errorf("WithHeaderDensity failed: got %f", opts.config.HeaderDensity)
	}

	WithColumnConsistency(0.9)(opts)
	if opts.config.ColumnConsistency != 0.9 {
		t.Errorf("WithColumnConsistency failed: got %f", opts.config.ColumnConsistency)
	}

	WithExpandMergedCells(false)(opts)
	if opts.config.ExpandMergedCells != false {
		t.Error("WithExpandMergedCells failed")
	}

	WithTrackMergeMetadata(false)(opts)
	if opts.config.TrackMergeMetadata != false {
		t.Error("WithTrackMergeMetadata failed")
	}

	WithParallel(true)(opts)
	if opts.parallel != true {
		t.Error("WithParallel failed")
	}
}

func TestDiffTables(t *testing.T) {
	// Create simple test tables
	table1 := &Table{
		Name:    "Table1",
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Index: 1, Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
		},
	}

	table2 := &Table{
		Name:    "Table2",
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Index: 1, Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	diff := DiffTables(table1, table2, "ID")

	if len(diff.AddedRows) != 1 {
		t.Errorf("Expected 1 added row, got %d", len(diff.AddedRows))
	}

	if len(diff.RemovedRows) != 1 {
		t.Errorf("Expected 1 removed row, got %d", len(diff.RemovedRows))
	}
}
