package reader

import (
	"path/filepath"
	"testing"

	"excel-lite/pkg/models"

	"github.com/xuri/excelize/v2"
)

// =============================================================================
// Benchmark Helpers
// =============================================================================

func createBenchmarkFile(b *testing.B, rows, cols int) string {
	b.Helper()
	dir := b.TempDir()
	path := filepath.Join(dir, "benchmark.xlsx")

	f := excelize.NewFile()

	// Create headers
	for col := 0; col < cols; col++ {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue("Sheet1", cell, "Header"+string(rune('A'+col)))
	}

	// Create data rows
	for row := 2; row <= rows+1; row++ {
		for col := 0; col < cols; col++ {
			cell, _ := excelize.CoordinatesToCellName(col+1, row)
			switch col % 4 {
			case 0:
				f.SetCellValue("Sheet1", cell, "String Value")
			case 1:
				f.SetCellValue("Sheet1", cell, float64(row*col))
			case 2:
				f.SetCellValue("Sheet1", cell, row%2 == 0)
			case 3:
				f.SetCellValue("Sheet1", cell, "2025-01-01")
			}
		}
	}

	if err := f.SaveAs(path); err != nil {
		b.Fatalf("Failed to create benchmark file: %v", err)
	}

	return path
}

func createBenchmarkFileWithMerges(b *testing.B, rows, cols int) string {
	b.Helper()
	dir := b.TempDir()
	path := filepath.Join(dir, "benchmark_merge.xlsx")

	f := excelize.NewFile()

	// Create merged header row
	f.SetCellValue("Sheet1", "A1", "Group A")
	f.MergeCell("Sheet1", "A1", "B1")
	f.SetCellValue("Sheet1", "C1", "Group B")
	f.MergeCell("Sheet1", "C1", "D1")

	// Create sub-headers
	f.SetCellValue("Sheet1", "A2", "Sub1")
	f.SetCellValue("Sheet1", "B2", "Sub2")
	f.SetCellValue("Sheet1", "C2", "Sub3")
	f.SetCellValue("Sheet1", "D2", "Sub4")

	// Create data rows
	for row := 3; row <= rows+2; row++ {
		for col := 0; col < cols && col < 4; col++ {
			cell, _ := excelize.CoordinatesToCellName(col+1, row)
			f.SetCellValue("Sheet1", cell, float64(row*col))
		}
	}

	if err := f.SaveAs(path); err != nil {
		b.Fatalf("Failed to create benchmark file: %v", err)
	}

	return path
}

// =============================================================================
// File Loading Benchmarks
// =============================================================================

func BenchmarkLoadFile_Small(b *testing.B) {
	path := createBenchmarkFile(b, 100, 5)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ef, err := LoadFile(path)
		if err != nil {
			b.Fatal(err)
		}
		ef.Close()
	}
}

func BenchmarkLoadFile_Medium(b *testing.B) {
	path := createBenchmarkFile(b, 1000, 10)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ef, err := LoadFile(path)
		if err != nil {
			b.Fatal(err)
		}
		ef.Close()
	}
}

func BenchmarkLoadFile_Large(b *testing.B) {
	path := createBenchmarkFile(b, 10000, 20)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ef, err := LoadFile(path)
		if err != nil {
			b.Fatal(err)
		}
		ef.Close()
	}
}

// =============================================================================
// Sheet Reading Benchmarks
// =============================================================================

func BenchmarkReadSheet_Small(b *testing.B) {
	path := createBenchmarkFile(b, 100, 5)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := sp.ReadSheet("Sheet1")
		if err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkReadSheet_Medium(b *testing.B) {
	path := createBenchmarkFile(b, 1000, 10)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := sp.ReadSheet("Sheet1")
		if err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkReadSheet_Large(b *testing.B) {
	path := createBenchmarkFile(b, 10000, 20)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := sp.ReadSheet("Sheet1")
		if err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkReadSheet_WithMergedCells(b *testing.B) {
	path := createBenchmarkFileWithMerges(b, 1000, 4)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := sp.ReadSheet("Sheet1")
		if err != nil {
			b.Fatal(err)
		}
	}
}

// =============================================================================
// Table Detection Benchmarks
// =============================================================================

func BenchmarkTableDetection_Small(b *testing.B) {
	path := createBenchmarkFile(b, 100, 5)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, _ := sp.ReadSheet("Sheet1")
	ta := NewDefaultAnalyzer()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ta.DetectTables(grid)
	}
}

func BenchmarkTableDetection_Medium(b *testing.B) {
	path := createBenchmarkFile(b, 1000, 10)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, _ := sp.ReadSheet("Sheet1")
	ta := NewDefaultAnalyzer()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ta.DetectTables(grid)
	}
}

func BenchmarkTableDetection_Large(b *testing.B) {
	path := createBenchmarkFile(b, 10000, 20)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, _ := sp.ReadSheet("Sheet1")
	ta := NewDefaultAnalyzer()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ta.DetectTables(grid)
	}
}

// =============================================================================
// Header Detection Benchmarks
// =============================================================================

func BenchmarkHeaderDetection(b *testing.B) {
	path := createBenchmarkFile(b, 100, 10)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, _ := sp.ReadSheet("Sheet1")
	ta := NewDefaultAnalyzer()
	boundaries := ta.DetectTables(grid)
	hd := NewDefaultHeaderDetector()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		for _, boundary := range boundaries {
			hd.DetectHeaderRow(grid, boundary)
		}
	}
}

func BenchmarkHeaderExtraction(b *testing.B) {
	path := createBenchmarkFile(b, 100, 20)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, _ := sp.ReadSheet("Sheet1")
	ta := NewDefaultAnalyzer()
	boundaries := ta.DetectTables(grid)
	hd := NewDefaultHeaderDetector()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		for _, boundary := range boundaries {
			headerRow := hd.DetectHeaderRow(grid, boundary)
			hd.ExtractHeaders(grid, headerRow, boundary)
		}
	}
}

// =============================================================================
// Full Workbook Reading Benchmarks
// =============================================================================

func BenchmarkWorkbookReader_Small(b *testing.B) {
	path := createBenchmarkFile(b, 100, 5)
	wr := NewWorkbookReader()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := wr.ReadFile(path)
		if err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkWorkbookReader_Medium(b *testing.B) {
	path := createBenchmarkFile(b, 1000, 10)
	wr := NewWorkbookReader()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := wr.ReadFile(path)
		if err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkWorkbookReader_Large(b *testing.B) {
	path := createBenchmarkFile(b, 10000, 20)
	wr := NewWorkbookReader()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := wr.ReadFile(path)
		if err != nil {
			b.Fatal(err)
		}
	}
}

// =============================================================================
// Row Parsing Benchmarks
// =============================================================================

func BenchmarkRowParsing(b *testing.B) {
	path := createBenchmarkFile(b, 1000, 10)
	ef, _ := LoadFile(path)
	defer ef.Close()

	sp := NewSheetProcessor(ef)
	grid, _ := sp.ReadSheet("Sheet1")
	ta := NewDefaultAnalyzer()
	boundaries := ta.DetectTables(grid)
	hd := NewDefaultHeaderDetector()
	rp := NewRowParser(models.DefaultConfig())

	var headers []string
	var headerRow int
	var boundary models.TableBoundary
	if len(boundaries) > 0 {
		boundary = boundaries[0]
		headerRow = hd.DetectHeaderRow(grid, boundary)
		headers = hd.ExtractHeaders(grid, headerRow, boundary)
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		rp.ParseTable(grid, boundary, headers, headerRow, "Table1")
	}
}

// =============================================================================
// Merge Processing Benchmarks
// =============================================================================

func BenchmarkMergeProcessing(b *testing.B) {
	config := models.DefaultConfig()
	mp := NewMergeProcessor(config)

	// Create a grid with many cells
	grid := make([][]models.Cell, 100)
	for i := range grid {
		grid[i] = make([]models.Cell, 20)
		for j := range grid[i] {
			grid[i][j] = models.Cell{Row: i, Col: j, Type: models.CellTypeString, Value: "test"}
		}
	}

	// Create multiple merge ranges
	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 2, EndCol: 3, Value: "Merge1"},
		{StartRow: 5, StartCol: 5, EndRow: 7, EndCol: 10, Value: "Merge2"},
		{StartRow: 10, StartCol: 0, EndRow: 15, EndCol: 5, Value: "Merge3"},
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		mp.ApplyMerges(grid, merges)
	}
}
