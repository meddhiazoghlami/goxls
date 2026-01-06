package stream

import (
	"context"
	"io"
	"path/filepath"
	"testing"
	"time"

	"github.com/meddhiazoghlami/goxls/pkg/models"
	"github.com/xuri/excelize/v2"
)

// =============================================================================
// Test Helpers
// =============================================================================

func createTestFile(t *testing.T, setup func(*excelize.File)) string {
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

func TestNewStreamReader_Simple(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 30)
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	if sr == nil {
		t.Fatal("NewStreamReader() returned nil")
	}

	headers := sr.Headers()
	if len(headers) != 2 {
		t.Errorf("Headers() len = %d, want 2", len(headers))
	}

	if headers[0] != "Name" {
		t.Errorf("Headers()[0] = %q, want %q", headers[0], "Name")
	}

	if headers[1] != "Age" {
		t.Errorf("Headers()[1] = %q, want %q", headers[1], "Age")
	}
}

func TestNewStreamReader_FileNotFound(t *testing.T) {
	_, err := NewStreamReader("/nonexistent/path/file.xlsx", "Sheet1")
	if err == nil {
		t.Fatal("NewStreamReader() expected error for nonexistent file")
	}
}

func TestNewStreamReader_SheetNotFound(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Test")
	})

	_, err := NewStreamReader(path, "NonexistentSheet")
	if err == nil {
		t.Fatal("NewStreamReader() expected error for nonexistent sheet")
	}
}

func TestNewStreamReader_EmptySheet(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		// Empty sheet
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil && err != ErrNoHeaders {
		// Either we get an error or we get a reader
		return
	}
	if sr != nil {
		defer sr.Close()
	}
}

// =============================================================================
// Next() Tests
// =============================================================================

func TestStreamReader_Next_Basic(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", 30)
		f.SetCellValue("Sheet1", "A3", "Bob")
		f.SetCellValue("Sheet1", "B3", 25)
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	// First row
	row1, err := sr.Next()
	if err != nil {
		t.Fatalf("Next() error = %v", err)
	}

	if row1.Index != 0 {
		t.Errorf("row1.Index = %d, want 0", row1.Index)
	}

	nameCell, ok := row1.Values["Name"]
	if !ok {
		t.Fatal("row1.Values[Name] not found")
	}
	if nameCell.AsString() != "Alice" {
		t.Errorf("Name = %q, want %q", nameCell.AsString(), "Alice")
	}

	// Second row
	row2, err := sr.Next()
	if err != nil {
		t.Fatalf("Next() error = %v", err)
	}

	if row2.Index != 1 {
		t.Errorf("row2.Index = %d, want 1", row2.Index)
	}

	nameCell2, ok := row2.Values["Name"]
	if !ok {
		t.Fatal("row2.Values[Name] not found")
	}
	if nameCell2.AsString() != "Bob" {
		t.Errorf("Name = %q, want %q", nameCell2.AsString(), "Bob")
	}

	// EOF
	_, err = sr.Next()
	if err != io.EOF {
		t.Errorf("Next() error = %v, want io.EOF", err)
	}
}

func TestStreamReader_Next_ManyRows(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "ID")
		f.SetCellValue("Sheet1", "B1", "Value")
		for i := 2; i <= 101; i++ {
			f.SetCellValue("Sheet1", "A"+string(rune('0'+i%10)), i-1)
			cell := "A" + string(rune(48+(i/10)%10)) + string(rune(48+i%10))
			if i < 10 {
				cell = "A" + string(rune(48+i))
			} else {
				cell, _ = excelize.CoordinatesToCellName(1, i)
			}
			f.SetCellValue("Sheet1", cell, i-1)
			cell, _ = excelize.CoordinatesToCellName(2, i)
			f.SetCellValue("Sheet1", cell, (i-1)*10)
		}
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	count := 0
	for {
		_, err := sr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("Next() error = %v at row %d", err, count)
		}
		count++
	}

	if count != 100 {
		t.Errorf("Read %d rows, want 100", count)
	}

	if sr.TotalRowsRead() != 100 {
		t.Errorf("TotalRowsRead() = %d, want 100", sr.TotalRowsRead())
	}
}

func TestStreamReader_Next_ClosedStream(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "A2", "Alice")
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}

	sr.Close()

	_, err = sr.Next()
	if err != ErrStreamClosed {
		t.Errorf("Next() error = %v, want ErrStreamClosed", err)
	}
}

// =============================================================================
// Options Tests
// =============================================================================

func TestStreamReader_WithStreamHeaders(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		// No headers, just data
		f.SetCellValue("Sheet1", "A1", "Alice")
		f.SetCellValue("Sheet1", "B1", 30)
		f.SetCellValue("Sheet1", "A2", "Bob")
		f.SetCellValue("Sheet1", "B2", 25)
	})

	sr, err := NewStreamReader(path, "Sheet1",
		WithStreamHeaders("Name", "Age"),
	)
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	headers := sr.Headers()
	if len(headers) != 2 {
		t.Errorf("Headers() len = %d, want 2", len(headers))
	}

	if headers[0] != "Name" || headers[1] != "Age" {
		t.Errorf("Headers() = %v, want [Name Age]", headers)
	}

	// First row should be Alice (data, not headers)
	row, err := sr.Next()
	if err != nil {
		t.Fatalf("Next() error = %v", err)
	}

	nameCell, ok := row.Values["Name"]
	if !ok {
		t.Fatal("row.Values[Name] not found")
	}
	if nameCell.AsString() != "Alice" {
		t.Errorf("Name = %q, want %q", nameCell.AsString(), "Alice")
	}
}

func TestStreamReader_WithStreamNoHeaders(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Alice")
		f.SetCellValue("Sheet1", "B1", 30)
	})

	sr, err := NewStreamReader(path, "Sheet1",
		WithStreamNoHeaders(),
	)
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	row, err := sr.Next()
	if err != nil {
		t.Fatalf("Next() error = %v", err)
	}

	// Headers should be generated as Column_1, Column_2
	headers := sr.Headers()
	if headers[0] != "Column_1" || headers[1] != "Column_2" {
		t.Errorf("Headers() = %v, want [Column_1 Column_2]", headers)
	}

	// First row should be Alice
	cell, ok := row.Values["Column_1"]
	if !ok {
		t.Fatal("row.Values[Column_1] not found")
	}
	if cell.AsString() != "Alice" {
		t.Errorf("Column_1 = %q, want %q", cell.AsString(), "Alice")
	}
}

func TestStreamReader_WithStreamSkipRows(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Title Row")
		f.SetCellValue("Sheet1", "A2", "Subtitle")
		f.SetCellValue("Sheet1", "A3", "Name")
		f.SetCellValue("Sheet1", "B3", "Age")
		f.SetCellValue("Sheet1", "A4", "Alice")
		f.SetCellValue("Sheet1", "B4", 30)
	})

	sr, err := NewStreamReader(path, "Sheet1",
		WithStreamSkipRows(2), // Skip title and subtitle rows
	)
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	headers := sr.Headers()
	if headers[0] != "Name" {
		t.Errorf("Headers()[0] = %q, want %q", headers[0], "Name")
	}

	row, err := sr.Next()
	if err != nil {
		t.Fatalf("Next() error = %v", err)
	}

	nameCell, ok := row.Values["Name"]
	if !ok {
		t.Fatal("row.Values[Name] not found")
	}
	if nameCell.AsString() != "Alice" {
		t.Errorf("Name = %q, want %q", nameCell.AsString(), "Alice")
	}
}

func TestStreamReader_WithStreamSkipEmptyRows(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "A2", "Alice")
		// Row 3 is empty
		f.SetCellValue("Sheet1", "A4", "Bob")
	})

	// With skip empty rows (default)
	sr, err := NewStreamReader(path, "Sheet1",
		WithStreamSkipEmptyRows(true),
	)
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}

	count := 0
	for {
		_, err := sr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("Next() error = %v", err)
		}
		count++
	}
	sr.Close()

	if count != 2 {
		t.Errorf("With skip empty: got %d rows, want 2", count)
	}

	// Without skip empty rows
	sr2, err := NewStreamReader(path, "Sheet1",
		WithStreamSkipEmptyRows(false),
	)
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr2.Close()

	count2 := 0
	for {
		_, err := sr2.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("Next() error = %v", err)
		}
		count2++
	}

	if count2 != 3 {
		t.Errorf("Without skip empty: got %d rows, want 3", count2)
	}
}

func TestStreamReader_WithStreamTypeDetection(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Value")
		f.SetCellValue("Sheet1", "A2", "123")
		f.SetCellValue("Sheet1", "A3", "true")
		f.SetCellValue("Sheet1", "A4", "text")
	})

	// With type detection (default)
	sr, err := NewStreamReader(path, "Sheet1",
		WithStreamTypeDetection(true),
	)
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}

	row, _ := sr.Next()
	cell := row.Values["Value"]
	if cell.Type != models.CellTypeNumber {
		t.Errorf("Type = %v, want CellTypeNumber", cell.Type)
	}

	row, _ = sr.Next()
	cell = row.Values["Value"]
	if cell.Type != models.CellTypeBool {
		t.Errorf("Type = %v, want CellTypeBool", cell.Type)
	}

	row, _ = sr.Next()
	cell = row.Values["Value"]
	if cell.Type != models.CellTypeString {
		t.Errorf("Type = %v, want CellTypeString", cell.Type)
	}
	sr.Close()

	// Without type detection
	sr2, err := NewStreamReader(path, "Sheet1",
		WithStreamTypeDetection(false),
	)
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr2.Close()

	row, _ = sr2.Next()
	cell = row.Values["Value"]
	if cell.Type != models.CellTypeString {
		t.Errorf("Without detection: Type = %v, want CellTypeString", cell.Type)
	}
}

// =============================================================================
// Type Inference Tests
// =============================================================================

func TestStreamReader_TypeInference(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Number")
		f.SetCellValue("Sheet1", "B1", "Bool")
		f.SetCellValue("Sheet1", "C1", "Date")
		f.SetCellValue("Sheet1", "D1", "String")
		f.SetCellValue("Sheet1", "E1", "Empty")

		f.SetCellValue("Sheet1", "A2", "42.5")
		f.SetCellValue("Sheet1", "B2", "true")
		f.SetCellValue("Sheet1", "C2", "2024-01-15")
		f.SetCellValue("Sheet1", "D2", "hello")
		// E2 is empty
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	row, err := sr.Next()
	if err != nil {
		t.Fatalf("Next() error = %v", err)
	}

	tests := []struct {
		column   string
		wantType models.CellType
	}{
		{"Number", models.CellTypeNumber},
		{"Bool", models.CellTypeBool},
		{"Date", models.CellTypeDate},
		{"String", models.CellTypeString},
		{"Empty", models.CellTypeEmpty},
	}

	for _, tt := range tests {
		cell, ok := row.Values[tt.column]
		if !ok && tt.wantType != models.CellTypeEmpty {
			t.Errorf("column %q not found", tt.column)
			continue
		}
		if cell.Type != tt.wantType {
			t.Errorf("column %q: Type = %v, want %v", tt.column, cell.Type, tt.wantType)
		}
	}
}

// =============================================================================
// ForEach Tests
// =============================================================================

func TestStreamReader_ForEach(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "A3", "Bob")
		f.SetCellValue("Sheet1", "A4", "Charlie")
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	names := []string{}
	err = sr.ForEach(func(row *StreamRow) error {
		cell, _ := row.Values["Name"]
		names = append(names, cell.AsString())
		return nil
	})

	if err != nil {
		t.Fatalf("ForEach() error = %v", err)
	}

	if len(names) != 3 {
		t.Errorf("ForEach processed %d rows, want 3", len(names))
	}

	if names[0] != "Alice" || names[1] != "Bob" || names[2] != "Charlie" {
		t.Errorf("names = %v, want [Alice Bob Charlie]", names)
	}
}

// =============================================================================
// CollectN Tests
// =============================================================================

func TestStreamReader_CollectN(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		for i := 2; i <= 11; i++ {
			cell, _ := excelize.CoordinatesToCellName(1, i)
			f.SetCellValue("Sheet1", cell, i-1)
		}
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	rows, err := sr.CollectN(5)
	if err != nil {
		t.Fatalf("CollectN() error = %v", err)
	}

	if len(rows) != 5 {
		t.Errorf("CollectN(5) returned %d rows, want 5", len(rows))
	}
}

// =============================================================================
// Context Tests
// =============================================================================

func TestNewStreamReaderWithContext(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "A2", "Alice")
	})

	ctx := context.Background()
	sr, err := NewStreamReaderWithContext(ctx, path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReaderWithContext() error = %v", err)
	}
	defer sr.Close()

	row, err := sr.Next()
	if err != nil {
		t.Fatalf("Next() error = %v", err)
	}

	cell, _ := row.Values["Name"]
	if cell.AsString() != "Alice" {
		t.Errorf("Name = %q, want %q", cell.AsString(), "Alice")
	}
}

func TestNewStreamReaderWithContext_Canceled(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
	})

	ctx, cancel := context.WithCancel(context.Background())
	cancel() // Cancel immediately

	_, err := NewStreamReaderWithContext(ctx, path, "Sheet1")
	if err == nil {
		t.Fatal("NewStreamReaderWithContext() expected error for canceled context")
	}
}

func TestStreamReader_Next_ContextCanceled(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		for i := 2; i <= 100; i++ {
			cell, _ := excelize.CoordinatesToCellName(1, i)
			f.SetCellValue("Sheet1", cell, i)
		}
	})

	ctx, cancel := context.WithCancel(context.Background())
	sr, err := NewStreamReaderWithContext(ctx, path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReaderWithContext() error = %v", err)
	}
	defer sr.Close()

	// Read a few rows
	sr.Next()
	sr.Next()

	// Cancel context
	cancel()

	// Next call should fail
	_, err = sr.Next()
	if err == nil {
		t.Error("Next() expected error after context cancellation")
	}
}

// =============================================================================
// Close Tests
// =============================================================================

func TestStreamReader_Close(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "A2", "Alice")
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}

	// First close should succeed
	err = sr.Close()
	if err != nil {
		t.Errorf("Close() error = %v", err)
	}

	// Second close should be no-op
	err = sr.Close()
	if err != nil {
		t.Errorf("Second Close() error = %v", err)
	}
}

// =============================================================================
// Header Normalization Tests
// =============================================================================

func TestStreamReader_DuplicateHeaders(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "B1", "Name")
		f.SetCellValue("Sheet1", "C1", "Name")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", "Baker")
		f.SetCellValue("Sheet1", "C2", "Charlie")
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	headers := sr.Headers()
	if len(headers) != 3 {
		t.Fatalf("Headers() len = %d, want 3", len(headers))
	}

	// Should have Name, Name_2, Name_3
	if headers[0] != "Name" {
		t.Errorf("Headers()[0] = %q, want %q", headers[0], "Name")
	}
	if headers[1] != "Name_2" {
		t.Errorf("Headers()[1] = %q, want %q", headers[1], "Name_2")
	}
	if headers[2] != "Name_3" {
		t.Errorf("Headers()[2] = %q, want %q", headers[2], "Name_3")
	}
}

func TestStreamReader_EmptyHeaders(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		// B1 is empty
		f.SetCellValue("Sheet1", "C1", "Age")
		f.SetCellValue("Sheet1", "A2", "Alice")
		f.SetCellValue("Sheet1", "B2", "data")
		f.SetCellValue("Sheet1", "C2", 30)
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	headers := sr.Headers()
	if headers[1] != "Column_2" {
		t.Errorf("Headers()[1] = %q, want %q", headers[1], "Column_2")
	}
}

// =============================================================================
// Metadata Tests
// =============================================================================

func TestStreamReader_Metadata(t *testing.T) {
	path := createTestFile(t, func(f *excelize.File) {
		f.SetCellValue("Sheet1", "A1", "Name")
		f.SetCellValue("Sheet1", "A2", "Alice")
	})

	sr, err := NewStreamReader(path, "Sheet1")
	if err != nil {
		t.Fatalf("NewStreamReader() error = %v", err)
	}
	defer sr.Close()

	if sr.SheetName() != "Sheet1" {
		t.Errorf("SheetName() = %q, want %q", sr.SheetName(), "Sheet1")
	}

	if sr.FilePath() != path {
		t.Errorf("FilePath() = %q, want %q", sr.FilePath(), path)
	}
}

// =============================================================================
// StreamRow Tests
// =============================================================================

func TestStreamRow_Get(t *testing.T) {
	row := &StreamRow{
		Index: 0,
		Values: map[string]StreamCell{
			"Name": {Value: "Alice", Type: models.CellTypeString, RawValue: "Alice"},
		},
	}

	cell, ok := row.Get("Name")
	if !ok {
		t.Fatal("Get(Name) returned false")
	}
	if cell.AsString() != "Alice" {
		t.Errorf("Get(Name).AsString() = %q, want %q", cell.AsString(), "Alice")
	}

	_, ok = row.Get("NonExistent")
	if ok {
		t.Error("Get(NonExistent) returned true, want false")
	}
}

func TestStreamRow_IsEmpty(t *testing.T) {
	emptyRow := &StreamRow{
		Cells: []StreamCell{
			{Type: models.CellTypeEmpty, RawValue: ""},
			{Type: models.CellTypeEmpty, RawValue: ""},
		},
	}

	if !emptyRow.IsEmpty() {
		t.Error("IsEmpty() = false for empty row, want true")
	}

	nonEmptyRow := &StreamRow{
		Cells: []StreamCell{
			{Type: models.CellTypeEmpty, RawValue: ""},
			{Type: models.CellTypeString, RawValue: "data"},
		},
	}

	if nonEmptyRow.IsEmpty() {
		t.Error("IsEmpty() = true for non-empty row, want false")
	}
}

// =============================================================================
// StreamCell Tests
// =============================================================================

func TestStreamCell_AsFloat(t *testing.T) {
	cell := StreamCell{Value: 42.5, Type: models.CellTypeNumber}

	val, ok := cell.AsFloat()
	if !ok {
		t.Error("AsFloat() ok = false, want true")
	}
	if val != 42.5 {
		t.Errorf("AsFloat() = %f, want 42.5", val)
	}

	stringCell := StreamCell{Value: "text", Type: models.CellTypeString}
	_, ok = stringCell.AsFloat()
	if ok {
		t.Error("AsFloat() ok = true for string, want false")
	}
}

// =============================================================================
// Type Inferrer Tests
// =============================================================================

func TestTypeInferrer_InferType(t *testing.T) {
	ti := NewTypeInferrer(nil)

	tests := []struct {
		value    string
		wantType models.CellType
	}{
		{"", models.CellTypeEmpty},
		{"hello", models.CellTypeString},
		{"123", models.CellTypeNumber},
		{"123.456", models.CellTypeNumber},
		{"-123.456", models.CellTypeNumber},
		{"1.5e10", models.CellTypeNumber},
		{"true", models.CellTypeBool},
		{"false", models.CellTypeBool},
		{"TRUE", models.CellTypeBool},
		{"FALSE", models.CellTypeBool},
		{"2024-01-15", models.CellTypeDate},
		{"01/15/2024", models.CellTypeDate},
		{"Jan 15, 2024", models.CellTypeDate},
	}

	for _, tt := range tests {
		got := ti.InferType(tt.value)
		if got != tt.wantType {
			t.Errorf("InferType(%q) = %v, want %v", tt.value, got, tt.wantType)
		}
	}
}

func TestTypeInferrer_ParseValue(t *testing.T) {
	ti := NewTypeInferrer(nil)

	// Number
	val := ti.ParseValue("42.5", models.CellTypeNumber)
	if f, ok := val.(float64); !ok || f != 42.5 {
		t.Errorf("ParseValue(42.5, Number) = %v, want 42.5", val)
	}

	// Bool
	val = ti.ParseValue("true", models.CellTypeBool)
	if b, ok := val.(bool); !ok || !b {
		t.Errorf("ParseValue(true, Bool) = %v, want true", val)
	}

	// Date
	val = ti.ParseValue("2024-01-15", models.CellTypeDate)
	if _, ok := val.(time.Time); !ok {
		t.Errorf("ParseValue(2024-01-15, Date) type = %T, want time.Time", val)
	}

	// String
	val = ti.ParseValue("hello", models.CellTypeString)
	if s, ok := val.(string); !ok || s != "hello" {
		t.Errorf("ParseValue(hello, String) = %v, want hello", val)
	}

	// Empty
	val = ti.ParseValue("", models.CellTypeEmpty)
	if val != nil {
		t.Errorf("ParseValue('', Empty) = %v, want nil", val)
	}
}

func TestTypeInferrer_CustomDateFormats(t *testing.T) {
	customFormats := []string{"02.01.2006"} // DD.MM.YYYY
	ti := NewTypeInferrer(customFormats)

	// Should recognize custom format
	if ti.InferType("15.01.2024") != models.CellTypeDate {
		t.Error("Custom format 15.01.2024 not recognized as date")
	}

	// Should NOT recognize standard format (since we overrode)
	if ti.InferType("2024-01-15") == models.CellTypeDate {
		t.Error("Standard format should not be recognized with custom formats")
	}
}
