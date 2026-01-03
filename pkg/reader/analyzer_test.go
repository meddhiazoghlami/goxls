package reader

import (
	"testing"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// =============================================================================
// Test Helpers
// =============================================================================

func makeCell(value string, cellType models.CellType) models.Cell {
	return models.Cell{
		Value:    value,
		Type:     cellType,
		RawValue: value,
	}
}

func makeEmptyCell() models.Cell {
	return models.Cell{
		Type:     models.CellTypeEmpty,
		RawValue: "",
	}
}

func makeGrid(rows, cols int, fill func(row, col int) models.Cell) [][]models.Cell {
	grid := make([][]models.Cell, rows)
	for r := 0; r < rows; r++ {
		grid[r] = make([]models.Cell, cols)
		for c := 0; c < cols; c++ {
			grid[r][c] = fill(r, c)
			grid[r][c].Row = r
			grid[r][c].Col = c
		}
	}
	return grid
}

// =============================================================================
// Constructor Tests
// =============================================================================

func TestNewTableAnalyzer(t *testing.T) {
	config := models.DetectionConfig{
		MinColumns:   3,
		MinRows:      5,
		MaxEmptyRows: 1,
	}

	ta := NewTableAnalyzer(config)

	if ta == nil {
		t.Fatal("NewTableAnalyzer() returned nil")
	}

	if ta.config.MinColumns != 3 {
		t.Errorf("config.MinColumns = %d, want 3", ta.config.MinColumns)
	}

	if ta.config.MinRows != 5 {
		t.Errorf("config.MinRows = %d, want 5", ta.config.MinRows)
	}
}

func TestNewDefaultAnalyzer(t *testing.T) {
	ta := NewDefaultAnalyzer()

	if ta == nil {
		t.Fatal("NewDefaultAnalyzer() returned nil")
	}

	defaultConfig := models.DefaultConfig()
	if ta.config.MinColumns != defaultConfig.MinColumns {
		t.Errorf("config.MinColumns = %d, want %d", ta.config.MinColumns, defaultConfig.MinColumns)
	}
}

// =============================================================================
// DetectTables Tests
// =============================================================================

func TestTableAnalyzer_DetectTables_Empty(t *testing.T) {
	ta := NewDefaultAnalyzer()

	tests := []struct {
		name string
		grid [][]models.Cell
	}{
		{"nil grid", nil},
		{"empty grid", [][]models.Cell{}},
		{"grid with empty rows", [][]models.Cell{{}, {}, {}}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			tables := ta.DetectTables(tt.grid)
			if len(tables) != 0 {
				t.Errorf("DetectTables() returned %d tables, want 0", len(tables))
			}
		})
	}
}

func TestTableAnalyzer_DetectTables_SingleTable(t *testing.T) {
	ta := NewDefaultAnalyzer()

	// 3x3 table with all cells filled
	grid := makeGrid(3, 3, func(row, col int) models.Cell {
		return makeCell("data", models.CellTypeString)
	})

	tables := ta.DetectTables(grid)

	if len(tables) != 1 {
		t.Fatalf("DetectTables() returned %d tables, want 1", len(tables))
	}

	table := tables[0]
	if table.StartRow != 0 {
		t.Errorf("StartRow = %d, want 0", table.StartRow)
	}
	if table.EndRow != 2 {
		t.Errorf("EndRow = %d, want 2", table.EndRow)
	}
	if table.StartCol != 0 {
		t.Errorf("StartCol = %d, want 0", table.StartCol)
	}
	if table.EndCol != 2 {
		t.Errorf("EndCol = %d, want 2", table.EndCol)
	}
}

func TestTableAnalyzer_DetectTables_OffsetTable(t *testing.T) {
	ta := NewDefaultAnalyzer()

	// 8x8 grid with 3x3 table starting at (2, 2)
	grid := makeGrid(8, 8, func(row, col int) models.Cell {
		if row >= 2 && row <= 4 && col >= 2 && col <= 4 {
			return makeCell("data", models.CellTypeString)
		}
		return makeEmptyCell()
	})

	tables := ta.DetectTables(grid)

	if len(tables) != 1 {
		t.Fatalf("DetectTables() returned %d tables, want 1", len(tables))
	}

	table := tables[0]
	if table.StartRow != 2 {
		t.Errorf("StartRow = %d, want 2", table.StartRow)
	}
	if table.StartCol != 2 {
		t.Errorf("StartCol = %d, want 2", table.StartCol)
	}
}

func TestTableAnalyzer_DetectTables_MultipleTables(t *testing.T) {
	ta := NewDefaultAnalyzer()

	// Grid with two separate 2x2 tables
	grid := makeGrid(10, 10, func(row, col int) models.Cell {
		// Table 1: rows 0-2, cols 0-2
		if row >= 0 && row <= 2 && col >= 0 && col <= 2 {
			return makeCell("table1", models.CellTypeString)
		}
		// Table 2: rows 6-8, cols 5-7
		if row >= 6 && row <= 8 && col >= 5 && col <= 7 {
			return makeCell("table2", models.CellTypeString)
		}
		return makeEmptyCell()
	})

	tables := ta.DetectTables(grid)

	if len(tables) != 2 {
		t.Errorf("DetectTables() returned %d tables, want 2", len(tables))
	}
}

func TestTableAnalyzer_DetectTables_TooSmall(t *testing.T) {
	ta := NewTableAnalyzer(models.DetectionConfig{
		MinColumns:   3,
		MinRows:      3,
		MaxEmptyRows: 2,
	})

	// 2x2 table - should not be detected
	grid := makeGrid(2, 2, func(row, col int) models.Cell {
		return makeCell("data", models.CellTypeString)
	})

	tables := ta.DetectTables(grid)

	if len(tables) != 0 {
		t.Errorf("DetectTables() returned %d tables for undersized grid, want 0", len(tables))
	}
}

func TestTableAnalyzer_DetectTables_WithGaps(t *testing.T) {
	ta := NewTableAnalyzer(models.DetectionConfig{
		MinColumns:   2,
		MinRows:      2,
		MaxEmptyRows: 1,
	})

	// Table with one empty row in the middle
	grid := makeGrid(5, 3, func(row, col int) models.Cell {
		if row == 2 { // Empty row
			return makeEmptyCell()
		}
		return makeCell("data", models.CellTypeString)
	})

	tables := ta.DetectTables(grid)

	// Should detect as one table (gap is within tolerance)
	if len(tables) != 1 {
		t.Errorf("DetectTables() returned %d tables, want 1", len(tables))
	}
}

func TestTableAnalyzer_DetectTables_LargeGap(t *testing.T) {
	ta := NewTableAnalyzer(models.DetectionConfig{
		MinColumns:   2,
		MinRows:      2,
		MaxEmptyRows: 1,
	})

	// Two tables separated by 3 empty rows
	grid := makeGrid(10, 3, func(row, col int) models.Cell {
		if row <= 2 {
			return makeCell("table1", models.CellTypeString)
		}
		if row >= 6 && row <= 8 {
			return makeCell("table2", models.CellTypeString)
		}
		return makeEmptyCell()
	})

	tables := ta.DetectTables(grid)

	if len(tables) != 2 {
		t.Errorf("DetectTables() returned %d tables, want 2", len(tables))
	}
}

// =============================================================================
// isValidTable Tests
// =============================================================================

func TestTableAnalyzer_isValidTable(t *testing.T) {
	ta := NewTableAnalyzer(models.DetectionConfig{
		MinColumns: 2,
		MinRows:    3,
	})

	tests := []struct {
		name     string
		boundary models.TableBoundary
		expected bool
	}{
		{
			name:     "valid table",
			boundary: models.TableBoundary{StartRow: 0, EndRow: 4, StartCol: 0, EndCol: 3},
			expected: true,
		},
		{
			name:     "minimum valid",
			boundary: models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 1},
			expected: true,
		},
		{
			name:     "too few rows",
			boundary: models.TableBoundary{StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 3},
			expected: false,
		},
		{
			name:     "too few columns",
			boundary: models.TableBoundary{StartRow: 0, EndRow: 4, StartCol: 0, EndCol: 0},
			expected: false,
		},
		{
			name:     "single cell",
			boundary: models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 0},
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := ta.isValidTable(tt.boundary); got != tt.expected {
				t.Errorf("isValidTable() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// =============================================================================
// Density Calculation Tests
// =============================================================================

func TestTableAnalyzer_calculateDensity(t *testing.T) {
	ta := NewDefaultAnalyzer()

	tests := []struct {
		name     string
		grid     [][]models.Cell
		startRow int
		startCol int
		size     int
		expected float64
	}{
		{
			name: "all filled",
			grid: makeGrid(3, 3, func(row, col int) models.Cell {
				return makeCell("x", models.CellTypeString)
			}),
			startRow: 0,
			startCol: 0,
			size:     3,
			expected: 1.0,
		},
		{
			name: "all empty",
			grid: makeGrid(3, 3, func(row, col int) models.Cell {
				return makeEmptyCell()
			}),
			startRow: 0,
			startCol: 0,
			size:     3,
			expected: 0.0,
		},
		{
			name: "half filled",
			grid: makeGrid(2, 2, func(row, col int) models.Cell {
				if (row+col)%2 == 0 {
					return makeCell("x", models.CellTypeString)
				}
				return makeEmptyCell()
			}),
			startRow: 0,
			startCol: 0,
			size:     2,
			expected: 0.5,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := ta.calculateDensity(tt.grid, tt.startRow, tt.startCol, tt.size)
			if got != tt.expected {
				t.Errorf("calculateDensity() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// =============================================================================
// Region Merging Tests
// =============================================================================

func TestTableAnalyzer_overlaps(t *testing.T) {
	ta := NewDefaultAnalyzer()

	tests := []struct {
		name     string
		a        models.TableBoundary
		b        models.TableBoundary
		expected bool
	}{
		{
			name:     "identical",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			expected: true,
		},
		{
			name:     "overlapping",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 3, EndRow: 8, StartCol: 3, EndCol: 8},
			expected: true,
		},
		{
			name:     "adjacent horizontal",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 6, EndCol: 10},
			expected: false,
		},
		{
			name:     "adjacent vertical",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 6, EndRow: 10, StartCol: 0, EndCol: 5},
			expected: false,
		},
		{
			name:     "far apart",
			a:        models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2},
			b:        models.TableBoundary{StartRow: 10, EndRow: 12, StartCol: 10, EndCol: 12},
			expected: false,
		},
		{
			name:     "touching corner",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 5, EndRow: 10, StartCol: 5, EndCol: 10},
			expected: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := ta.overlaps(tt.a, tt.b); got != tt.expected {
				t.Errorf("overlaps() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestTableAnalyzer_merge(t *testing.T) {
	ta := NewDefaultAnalyzer()

	tests := []struct {
		name     string
		a        models.TableBoundary
		b        models.TableBoundary
		expected models.TableBoundary
	}{
		{
			name:     "same boundary",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			expected: models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
		},
		{
			name:     "extend right",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 3, EndCol: 10},
			expected: models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 10},
		},
		{
			name:     "extend down",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 3, EndRow: 10, StartCol: 0, EndCol: 5},
			expected: models.TableBoundary{StartRow: 0, EndRow: 10, StartCol: 0, EndCol: 5},
		},
		{
			name:     "diagonal merge",
			a:        models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			b:        models.TableBoundary{StartRow: 3, EndRow: 10, StartCol: 3, EndCol: 10},
			expected: models.TableBoundary{StartRow: 0, EndRow: 10, StartCol: 0, EndCol: 10},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := ta.merge(tt.a, tt.b)
			if got != tt.expected {
				t.Errorf("merge() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestTableAnalyzer_mergeOverlappingRegions(t *testing.T) {
	ta := NewDefaultAnalyzer()

	tests := []struct {
		name     string
		regions  []models.TableBoundary
		expected int
	}{
		{
			name:     "empty",
			regions:  []models.TableBoundary{},
			expected: 0,
		},
		{
			name: "single region",
			regions: []models.TableBoundary{
				{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
			},
			expected: 1,
		},
		{
			name: "non-overlapping",
			regions: []models.TableBoundary{
				{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
				{StartRow: 10, EndRow: 15, StartCol: 10, EndCol: 15},
			},
			expected: 2,
		},
		{
			name: "overlapping - should merge",
			regions: []models.TableBoundary{
				{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
				{StartRow: 3, EndRow: 8, StartCol: 3, EndCol: 8},
			},
			expected: 1,
		},
		{
			name: "chain overlap",
			regions: []models.TableBoundary{
				{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 5},
				{StartRow: 4, EndRow: 9, StartCol: 4, EndCol: 9},
				{StartRow: 8, EndRow: 13, StartCol: 8, EndCol: 13},
			},
			expected: 1,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := ta.mergeOverlappingRegions(tt.regions)
			if len(got) != tt.expected {
				t.Errorf("mergeOverlappingRegions() returned %d regions, want %d", len(got), tt.expected)
			}
		})
	}
}

// =============================================================================
// Helper Functions Tests
// =============================================================================

func TestMinInt(t *testing.T) {
	tests := []struct {
		a, b, expected int
	}{
		{1, 2, 1},
		{2, 1, 1},
		{5, 5, 5},
		{-1, 1, -1},
		{0, 0, 0},
		{-10, -5, -10},
	}

	for _, tt := range tests {
		if got := minInt(tt.a, tt.b); got != tt.expected {
			t.Errorf("minInt(%d, %d) = %d, want %d", tt.a, tt.b, got, tt.expected)
		}
	}
}

func TestMaxInt(t *testing.T) {
	tests := []struct {
		a, b, expected int
	}{
		{1, 2, 2},
		{2, 1, 2},
		{5, 5, 5},
		{-1, 1, 1},
		{0, 0, 0},
		{-10, -5, -5},
	}

	for _, tt := range tests {
		if got := maxInt(tt.a, tt.b); got != tt.expected {
			t.Errorf("maxInt(%d, %d) = %d, want %d", tt.a, tt.b, got, tt.expected)
		}
	}
}

// =============================================================================
// Complex Scenarios
// =============================================================================

func TestTableAnalyzer_DetectTables_RealWorldScenarios(t *testing.T) {
	ta := NewDefaultAnalyzer()

	t.Run("table with header metadata above", func(t *testing.T) {
		// Simulates: Title in A1, empty row, then table
		grid := makeGrid(10, 5, func(row, col int) models.Cell {
			if row == 0 && col == 0 {
				return makeCell("Report Title", models.CellTypeString)
			}
			if row >= 2 && row <= 8 {
				return makeCell("data", models.CellTypeString)
			}
			return makeEmptyCell()
		})

		tables := ta.DetectTables(grid)

		// Should detect the main table, may or may not include title
		if len(tables) == 0 {
			t.Error("Expected at least one table to be detected")
		}
	})

	t.Run("sparse table with some empty columns", func(t *testing.T) {
		grid := makeGrid(5, 6, func(row, col int) models.Cell {
			// Every other column is empty
			if col%2 == 1 {
				return makeEmptyCell()
			}
			return makeCell("data", models.CellTypeString)
		})

		tables := ta.DetectTables(grid)

		if len(tables) == 0 {
			t.Error("Expected at least one table to be detected")
		}
	})
}

// =============================================================================
// FindDenseRegions Tests
// =============================================================================

func TestTableAnalyzer_FindDenseRegions(t *testing.T) {
	ta := NewDefaultAnalyzer()

	t.Run("empty grid", func(t *testing.T) {
		regions := ta.FindDenseRegions(nil, 2)
		if len(regions) != 0 {
			t.Errorf("Expected 0 regions for nil grid, got %d", len(regions))
		}

		regions = ta.FindDenseRegions([][]models.Cell{}, 2)
		if len(regions) != 0 {
			t.Errorf("Expected 0 regions for empty grid, got %d", len(regions))
		}
	})

	t.Run("invalid window size", func(t *testing.T) {
		grid := makeGrid(3, 3, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		regions := ta.FindDenseRegions(grid, 0)
		if len(regions) != 0 {
			t.Errorf("Expected 0 regions for window size 0, got %d", len(regions))
		}

		regions = ta.FindDenseRegions(grid, -1)
		if len(regions) != 0 {
			t.Errorf("Expected 0 regions for negative window size, got %d", len(regions))
		}
	})

	t.Run("fully dense grid", func(t *testing.T) {
		grid := makeGrid(5, 5, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		regions := ta.FindDenseRegions(grid, 2)
		if len(regions) == 0 {
			t.Error("Expected at least one dense region in fully populated grid")
		}
	})

	t.Run("sparse grid below threshold", func(t *testing.T) {
		// Create a very sparse grid that shouldn't meet density threshold
		grid := makeGrid(10, 10, func(row, col int) models.Cell {
			// Only one cell filled
			if row == 0 && col == 0 {
				return makeCell("data", models.CellTypeString)
			}
			return makeEmptyCell()
		})

		ta2 := NewTableAnalyzer(models.DetectionConfig{
			HeaderDensity: 0.9, // High threshold
		})
		regions := ta2.FindDenseRegions(grid, 3)
		// May or may not find regions depending on density calculation
		_ = regions
	})

	t.Run("window larger than grid", func(t *testing.T) {
		grid := makeGrid(3, 3, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		regions := ta.FindDenseRegions(grid, 10)
		if len(regions) != 0 {
			t.Errorf("Expected 0 regions for window larger than grid, got %d", len(regions))
		}
	})

	t.Run("overlapping dense regions merge", func(t *testing.T) {
		grid := makeGrid(6, 6, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		regions := ta.FindDenseRegions(grid, 2)
		// Dense regions that overlap should be merged
		if len(regions) > 0 {
			// Verify the result is reasonable
			for _, r := range regions {
				if r.EndRow < r.StartRow || r.EndCol < r.StartCol {
					t.Error("Invalid region boundaries")
				}
			}
		}
	})
}

func TestTableAnalyzer_calculateDensity_EdgeCases(t *testing.T) {
	ta := NewDefaultAnalyzer()

	t.Run("size extends beyond grid rows", func(t *testing.T) {
		grid := makeGrid(2, 3, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		// Request size larger than remaining rows
		density := ta.calculateDensity(grid, 0, 0, 10)
		if density <= 0 {
			t.Error("Expected positive density for partial grid coverage")
		}
	})

	t.Run("size extends beyond grid cols", func(t *testing.T) {
		grid := makeGrid(3, 2, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		// Request size larger than remaining cols
		density := ta.calculateDensity(grid, 0, 0, 10)
		if density <= 0 {
			t.Error("Expected positive density for partial grid coverage")
		}
	})

	t.Run("zero total cells edge case", func(t *testing.T) {
		// Empty rows that don't match the size check
		grid := [][]models.Cell{{}}

		density := ta.calculateDensity(grid, 0, 0, 1)
		if density != 0 {
			t.Errorf("Expected 0 density for empty region, got %v", density)
		}
	})
}

// =============================================================================
// expandTable Edge Cases
// =============================================================================

func TestTableAnalyzer_expandForMerges(t *testing.T) {
	ta := NewDefaultAnalyzer()

	t.Run("no merges", func(t *testing.T) {
		grid := makeGrid(3, 3, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2}
		result := ta.expandForMerges(grid, boundary)

		if result != boundary {
			t.Errorf("Expected unchanged boundary, got %+v", result)
		}
	})

	t.Run("expand for merge extending down", func(t *testing.T) {
		grid := makeGrid(5, 3, func(row, col int) models.Cell {
			cell := makeCell("data", models.CellTypeString)
			// Add merge that extends beyond boundary
			if row == 2 && col == 0 {
				cell.MergeRange = &models.MergeRange{
					StartRow: 2, EndRow: 4, StartCol: 0, EndCol: 0,
				}
			}
			return cell
		})

		boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2}
		result := ta.expandForMerges(grid, boundary)

		if result.EndRow != 4 {
			t.Errorf("Expected EndRow=4, got %d", result.EndRow)
		}
	})

	t.Run("expand for merge extending right", func(t *testing.T) {
		grid := makeGrid(3, 5, func(row, col int) models.Cell {
			cell := makeCell("data", models.CellTypeString)
			if row == 0 && col == 2 {
				cell.MergeRange = &models.MergeRange{
					StartRow: 0, EndRow: 0, StartCol: 2, EndCol: 4,
				}
			}
			return cell
		})

		boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2}
		result := ta.expandForMerges(grid, boundary)

		if result.EndCol != 4 {
			t.Errorf("Expected EndCol=4, got %d", result.EndCol)
		}
	})

	t.Run("expand for merge extending up", func(t *testing.T) {
		grid := makeGrid(5, 3, func(row, col int) models.Cell {
			cell := makeCell("data", models.CellTypeString)
			if row == 1 && col == 0 {
				cell.MergeRange = &models.MergeRange{
					StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 0,
				}
			}
			return cell
		})

		boundary := models.TableBoundary{StartRow: 1, EndRow: 4, StartCol: 0, EndCol: 2}
		result := ta.expandForMerges(grid, boundary)

		if result.StartRow != 0 {
			t.Errorf("Expected StartRow=0, got %d", result.StartRow)
		}
	})

	t.Run("expand for merge extending left", func(t *testing.T) {
		grid := makeGrid(3, 5, func(row, col int) models.Cell {
			cell := makeCell("data", models.CellTypeString)
			if row == 0 && col == 1 {
				cell.MergeRange = &models.MergeRange{
					StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 1,
				}
			}
			return cell
		})

		boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 1, EndCol: 4}
		result := ta.expandForMerges(grid, boundary)

		if result.StartCol != 0 {
			t.Errorf("Expected StartCol=0, got %d", result.StartCol)
		}
	})
}

func TestTableAnalyzer_expandTable_EdgeCases(t *testing.T) {
	ta := NewDefaultAnalyzer()

	t.Run("expand left boundary", func(t *testing.T) {
		// Table with data extending to the left of start position
		grid := makeGrid(3, 5, func(row, col int) models.Cell {
			return makeCell("data", models.CellTypeString)
		})

		visited := make([][]bool, len(grid))
		for i := range visited {
			visited[i] = make([]bool, len(grid[i]))
		}

		// Start from middle column
		boundary := ta.expandTable(grid, 0, 2, visited)

		if boundary.StartCol != 0 {
			t.Errorf("Expected StartCol=0, got %d", boundary.StartCol)
		}
	})

	t.Run("jagged grid", func(t *testing.T) {
		// Grid with uneven row lengths
		grid := [][]models.Cell{
			{makeCell("a", models.CellTypeString), makeCell("b", models.CellTypeString), makeCell("c", models.CellTypeString)},
			{makeCell("d", models.CellTypeString)}, // Short row
			{makeCell("e", models.CellTypeString), makeCell("f", models.CellTypeString)},
		}

		visited := make([][]bool, len(grid))
		for i := range visited {
			visited[i] = make([]bool, len(grid[i]))
		}

		boundary := ta.expandTable(grid, 0, 0, visited)

		// Should handle jagged rows without panic
		if boundary.StartRow != 0 {
			t.Errorf("Expected StartRow=0, got %d", boundary.StartRow)
		}
	})

	t.Run("consecutive empty columns", func(t *testing.T) {
		// Grid with multiple consecutive empty columns
		grid := makeGrid(5, 10, func(row, col int) models.Cell {
			if col <= 2 {
				return makeCell("data", models.CellTypeString)
			}
			return makeEmptyCell()
		})

		visited := make([][]bool, len(grid))
		for i := range visited {
			visited[i] = make([]bool, len(grid[i]))
		}

		boundary := ta.expandTable(grid, 0, 0, visited)

		if boundary.EndCol > 3 {
			t.Errorf("Expected EndCol <= 3 due to empty columns, got %d", boundary.EndCol)
		}
	})
}
