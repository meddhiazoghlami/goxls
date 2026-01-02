package reader

import (
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// =============================================================================
// Test Helpers
// =============================================================================

func createTestExcelFile(t *testing.T, filename string) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, filename)

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Test")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}
	return path
}

func createEmptyFile(t *testing.T, filename string) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, filename)

	f, err := os.Create(path)
	if err != nil {
		t.Fatalf("Failed to create empty file: %v", err)
	}
	f.Close()
	return path
}

// =============================================================================
// LoadFile Tests
// =============================================================================

func TestLoadFile_Success(t *testing.T) {
	path := createTestExcelFile(t, "valid.xlsx")

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v, want nil", err)
	}
	defer ef.Close()

	if ef.file == nil {
		t.Error("LoadFile() returned nil file handle")
	}

	if ef.FilePath() != path {
		t.Errorf("FilePath() = %v, want %v", ef.FilePath(), path)
	}
}

func TestLoadFile_FileNotFound(t *testing.T) {
	_, err := LoadFile("/nonexistent/path/file.xlsx")

	if err == nil {
		t.Fatal("LoadFile() expected error, got nil")
	}

	if !errors.Is(err, ErrFileNotFound) {
		t.Errorf("LoadFile() error = %v, want ErrFileNotFound", err)
	}
}

func TestLoadFile_InvalidExtension(t *testing.T) {
	tests := []struct {
		name     string
		filename string
	}{
		{"csv file", "test.csv"},
		{"xls file", "test.xls"},
		{"txt file", "test.txt"},
		{"no extension", "testfile"},
		{"xlsm file", "test.xlsm"},
		{"json file", "test.json"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dir := t.TempDir()
			path := filepath.Join(dir, tt.filename)

			// Create a file with content
			if err := os.WriteFile(path, []byte("test content"), 0644); err != nil {
				t.Fatalf("Failed to create test file: %v", err)
			}

			_, err := LoadFile(path)

			if err == nil {
				t.Fatal("LoadFile() expected error, got nil")
			}

			if !errors.Is(err, ErrInvalidFormat) {
				t.Errorf("LoadFile() error = %v, want ErrInvalidFormat", err)
			}
		})
	}
}

func TestLoadFile_EmptyFile(t *testing.T) {
	path := createEmptyFile(t, "empty.xlsx")

	_, err := LoadFile(path)

	if err == nil {
		t.Fatal("LoadFile() expected error, got nil")
	}

	if !errors.Is(err, ErrFileEmpty) {
		t.Errorf("LoadFile() error = %v, want ErrFileEmpty", err)
	}
}

func TestLoadFile_CorruptedFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "corrupted.xlsx")

	// Create a file with invalid xlsx content
	if err := os.WriteFile(path, []byte("not a valid xlsx file"), 0644); err != nil {
		t.Fatalf("Failed to create corrupted file: %v", err)
	}

	_, err := LoadFile(path)

	if err == nil {
		t.Fatal("LoadFile() expected error for corrupted file, got nil")
	}

	if !errors.Is(err, ErrCannotOpenFile) {
		t.Errorf("LoadFile() error = %v, want ErrCannotOpenFile", err)
	}
}

// =============================================================================
// ExcelFile Methods Tests
// =============================================================================

func TestExcelFile_Close(t *testing.T) {
	path := createTestExcelFile(t, "close_test.xlsx")

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}

	// Close should not return error
	if err := ef.Close(); err != nil {
		t.Errorf("Close() error = %v, want nil", err)
	}

	// Close on nil file should be safe
	ef2 := &ExcelFile{file: nil}
	if err := ef2.Close(); err != nil {
		t.Errorf("Close() on nil file error = %v, want nil", err)
	}
}

func TestExcelFile_GetSheetNames(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "multi_sheet.xlsx")

	f := excelize.NewFile()
	f.SetSheetName("Sheet1", "First")
	f.NewSheet("Second")
	f.NewSheet("Third")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	names := ef.GetSheetNames()

	if len(names) != 3 {
		t.Errorf("GetSheetNames() returned %d sheets, want 3", len(names))
	}

	expected := map[string]bool{"First": true, "Second": true, "Third": true}
	for _, name := range names {
		if !expected[name] {
			t.Errorf("Unexpected sheet name: %s", name)
		}
	}
}

func TestExcelFile_GetSheetCount(t *testing.T) {
	tests := []struct {
		name       string
		sheetNames []string
		expected   int
	}{
		{
			name:       "single sheet",
			sheetNames: []string{"Sheet1"},
			expected:   1,
		},
		{
			name:       "three sheets",
			sheetNames: []string{"A", "B", "C"},
			expected:   3,
		},
		{
			name:       "five sheets",
			sheetNames: []string{"S1", "S2", "S3", "S4", "S5"},
			expected:   5,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dir := t.TempDir()
			path := filepath.Join(dir, "test.xlsx")

			f := excelize.NewFile()
			for i, name := range tt.sheetNames {
				if i == 0 {
					f.SetSheetName("Sheet1", name)
				} else {
					f.NewSheet(name)
				}
			}
			if err := f.SaveAs(path); err != nil {
				t.Fatalf("Failed to create test file: %v", err)
			}

			ef, err := LoadFile(path)
			if err != nil {
				t.Fatalf("LoadFile() error = %v", err)
			}
			defer ef.Close()

			if got := ef.GetSheetCount(); got != tt.expected {
				t.Errorf("GetSheetCount() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestExcelFile_GetRows(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "rows_test.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Header1")
	f.SetCellValue("Sheet1", "B1", "Header2")
	f.SetCellValue("Sheet1", "A2", "Value1")
	f.SetCellValue("Sheet1", "B2", "Value2")
	f.SetCellValue("Sheet1", "A3", "Value3")
	f.SetCellValue("Sheet1", "B3", "Value4")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	rows, err := ef.GetRows("Sheet1")
	if err != nil {
		t.Fatalf("GetRows() error = %v", err)
	}

	if len(rows) != 3 {
		t.Errorf("GetRows() returned %d rows, want 3", len(rows))
	}

	if len(rows[0]) != 2 {
		t.Errorf("GetRows() first row has %d cols, want 2", len(rows[0]))
	}

	if rows[0][0] != "Header1" {
		t.Errorf("rows[0][0] = %v, want Header1", rows[0][0])
	}
}

func TestExcelFile_GetRows_NonexistentSheet(t *testing.T) {
	path := createTestExcelFile(t, "test.xlsx")

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	_, err = ef.GetRows("NonexistentSheet")
	if err == nil {
		t.Error("GetRows() expected error for nonexistent sheet, got nil")
	}
}

func TestExcelFile_GetCellValue(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "cell_test.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Hello")
	f.SetCellValue("Sheet1", "B2", 42)
	f.SetCellValue("Sheet1", "C3", 3.14159)
	f.SetCellValue("Sheet1", "D4", true)
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	tests := []struct {
		cell     string
		expected string
	}{
		{"A1", "Hello"},
		{"B2", "42"},
		{"C3", "3.14159"},
		{"D4", "TRUE"},
		{"E5", ""}, // Empty cell
	}

	for _, tt := range tests {
		t.Run(tt.cell, func(t *testing.T) {
			value, err := ef.GetCellValue("Sheet1", tt.cell)
			if err != nil {
				t.Fatalf("GetCellValue(%s) error = %v", tt.cell, err)
			}
			if value != tt.expected {
				t.Errorf("GetCellValue(%s) = %v, want %v", tt.cell, value, tt.expected)
			}
		})
	}
}

func TestExcelFile_GetCellType(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "type_test.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "String")
	f.SetCellValue("Sheet1", "B1", 42)
	f.SetCellValue("Sheet1", "C1", true)
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	// Test string cell
	ct, err := ef.GetCellType("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellType(A1) error = %v", err)
	}
	if ct != excelize.CellTypeSharedString && ct != excelize.CellTypeInlineString {
		t.Errorf("GetCellType(A1) = %v, want string type", ct)
	}

	// Test number cell - excelize may return different types based on how the cell was set
	ct, err = ef.GetCellType("Sheet1", "B1")
	if err != nil {
		t.Fatalf("GetCellType(B1) error = %v", err)
	}
	// Accept either Number type or Unset (excelize quirk with SetCellValue)
	if ct != excelize.CellTypeNumber && ct != excelize.CellTypeUnset {
		t.Errorf("GetCellType(B1) = %v, want CellTypeNumber or CellTypeUnset", ct)
	}
}

func TestExcelFile_Raw(t *testing.T) {
	path := createTestExcelFile(t, "raw_test.xlsx")

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	raw := ef.Raw()
	if raw == nil {
		t.Error("Raw() returned nil")
	}

	// Verify we can use the raw file
	names := raw.GetSheetList()
	if len(names) == 0 {
		t.Error("Raw file has no sheets")
	}
}

// =============================================================================
// Edge Cases
// =============================================================================

func TestLoadFile_SpecialCharactersInPath(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "file with spaces.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Test")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v for path with spaces", err)
	}
	defer ef.Close()
}

func TestLoadFile_UnicodeFilename(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "테스트파일.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Test")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v for unicode filename", err)
	}
	defer ef.Close()
}

func TestExcelFile_LargeFile(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping large file test in short mode")
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "large.xlsx")

	f := excelize.NewFile()
	// Create 1000 rows x 50 columns
	for row := 1; row <= 1000; row++ {
		for col := 1; col <= 50; col++ {
			cell, _ := excelize.CoordinatesToCellName(col, row)
			f.SetCellValue("Sheet1", cell, row*col)
		}
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create large test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v for large file", err)
	}
	defer ef.Close()

	rows, err := ef.GetRows("Sheet1")
	if err != nil {
		t.Fatalf("GetRows() error = %v", err)
	}

	if len(rows) != 1000 {
		t.Errorf("GetRows() returned %d rows, want 1000", len(rows))
	}
}

func TestExcelFile_GetCellFormula(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "formula_test.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "B1", 20)
	f.SetCellFormula("Sheet1", "C1", "A1+B1")
	f.SetCellFormula("Sheet1", "D1", "SUM(A1:B1)")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	tests := []struct {
		cell     string
		expected string
	}{
		{"A1", ""},         // Not a formula
		{"B1", ""},         // Not a formula
		{"C1", "A1+B1"},    // Formula
		{"D1", "SUM(A1:B1)"}, // Formula
	}

	for _, tt := range tests {
		t.Run(tt.cell, func(t *testing.T) {
			formula, err := ef.GetCellFormula("Sheet1", tt.cell)
			if err != nil {
				t.Fatalf("GetCellFormula(%s) error = %v", tt.cell, err)
			}
			if formula != tt.expected {
				t.Errorf("GetCellFormula(%s) = %q, want %q", tt.cell, formula, tt.expected)
			}
		})
	}
}

// =============================================================================
// Cell Comments Tests
// =============================================================================

func TestExcelFile_GetComments(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "comments_test.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Header")
	f.SetCellValue("Sheet1", "B1", "Value")
	f.AddComment("Sheet1", excelize.Comment{
		Cell:   "A1",
		Author: "TestAuthor",
		Text:   "This is a test comment",
	})
	f.AddComment("Sheet1", excelize.Comment{
		Cell:   "B1",
		Author: "Another Author",
		Text:   "Another comment",
	})
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	comments, err := ef.GetComments("Sheet1")
	if err != nil {
		t.Fatalf("GetComments() error = %v", err)
	}

	if len(comments) != 2 {
		t.Errorf("GetComments() returned %d comments, want 2", len(comments))
	}

	// Check first comment
	found := false
	for _, c := range comments {
		if c.Cell == "A1" {
			found = true
			if c.Author != "TestAuthor" {
				t.Errorf("Comment A1 Author = %q, want %q", c.Author, "TestAuthor")
			}
			if c.Text != "This is a test comment" {
				t.Errorf("Comment A1 Text = %q, want %q", c.Text, "This is a test comment")
			}
		}
	}
	if !found {
		t.Error("Comment for cell A1 not found")
	}
}

func TestExcelFile_GetComments_EmptySheet(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "no_comments.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "No comments here")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	comments, err := ef.GetComments("Sheet1")
	if err != nil {
		t.Fatalf("GetComments() error = %v", err)
	}

	if len(comments) != 0 {
		t.Errorf("GetComments() returned %d comments, want 0", len(comments))
	}
}

func TestExcelFile_GetComments_NonexistentSheet(t *testing.T) {
	path := createTestExcelFile(t, "test.xlsx")

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	_, err = ef.GetComments("NonexistentSheet")
	if err == nil {
		t.Error("GetComments() expected error for nonexistent sheet, got nil")
	}
}

func TestExcelFile_GetComments_UnicodeContent(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "unicode_comments.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Data")
	f.AddComment("Sheet1", excelize.Comment{
		Cell:   "A1",
		Author: "国际作者",
		Text:   "日本語コメント 한국어 中文",
	})
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	comments, err := ef.GetComments("Sheet1")
	if err != nil {
		t.Fatalf("GetComments() error = %v", err)
	}

	if len(comments) != 1 {
		t.Fatalf("GetComments() returned %d comments, want 1", len(comments))
	}

	if comments[0].Author != "国际作者" {
		t.Errorf("Comment Author = %q, want %q", comments[0].Author, "国际作者")
	}
	if comments[0].Text != "日本語コメント 한국어 中文" {
		t.Errorf("Comment Text = %q, want %q", comments[0].Text, "日本語コメント 한국어 中文")
	}
}

// =============================================================================
// Hyperlink Tests
// =============================================================================

func TestExcelFile_GetCellHyperLink(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "hyperlink_test.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Click here")
	f.SetCellValue("Sheet1", "B1", "No link")
	f.SetCellHyperLink("Sheet1", "A1", "https://example.com", "External")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	// Test cell with hyperlink
	link, err := ef.GetCellHyperLink("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellHyperLink(A1) error = %v", err)
	}
	if link != "https://example.com" {
		t.Errorf("GetCellHyperLink(A1) = %q, want %q", link, "https://example.com")
	}

	// Test cell without hyperlink
	link, err = ef.GetCellHyperLink("Sheet1", "B1")
	if err != nil {
		t.Fatalf("GetCellHyperLink(B1) error = %v", err)
	}
	if link != "" {
		t.Errorf("GetCellHyperLink(B1) = %q, want empty string", link)
	}
}

func TestExcelFile_GetCellHyperLink_NoLinks(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "no_links.xlsx")

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Plain text")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	link, err := ef.GetCellHyperLink("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellHyperLink() error = %v", err)
	}
	if link != "" {
		t.Errorf("GetCellHyperLink() = %q, want empty string", link)
	}
}

func TestExcelFile_GetCellHyperLink_NonexistentSheet(t *testing.T) {
	path := createTestExcelFile(t, "test.xlsx")

	ef, err := LoadFile(path)
	if err != nil {
		t.Fatalf("LoadFile() error = %v", err)
	}
	defer ef.Close()

	_, err = ef.GetCellHyperLink("NonexistentSheet", "A1")
	if err == nil {
		t.Error("GetCellHyperLink() expected error for nonexistent sheet, got nil")
	}
}
