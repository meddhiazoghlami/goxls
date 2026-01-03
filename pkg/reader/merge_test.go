package reader

import (
	"github.com/meddhiazoghlami/goxls/pkg/models"
	"testing"
)

func TestNewMergeProcessor(t *testing.T) {
	config := models.DefaultConfig()
	mp := NewMergeProcessor(config)

	if mp == nil {
		t.Fatal("NewMergeProcessor returned nil")
	}
	if mp.config.ExpandMergedCells != true {
		t.Error("Expected ExpandMergedCells to be true by default")
	}
	if mp.config.TrackMergeMetadata != true {
		t.Error("Expected TrackMergeMetadata to be true by default")
	}
}

func TestMergeProcessor_ParseMergeCells(t *testing.T) {
	tests := []struct {
		name     string
		merges   []MergeCellInfo
		expected []ParsedMergeRange
		wantErr  bool
	}{
		{
			name: "simple 2x2 merge",
			merges: []MergeCellInfo{
				{StartCell: "A1", EndCell: "B2", Value: "Merged"},
			},
			expected: []ParsedMergeRange{
				{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "Merged"},
			},
		},
		{
			name: "horizontal merge",
			merges: []MergeCellInfo{
				{StartCell: "A1", EndCell: "D1", Value: "Header"},
			},
			expected: []ParsedMergeRange{
				{StartRow: 0, StartCol: 0, EndRow: 0, EndCol: 3, Value: "Header"},
			},
		},
		{
			name: "vertical merge",
			merges: []MergeCellInfo{
				{StartCell: "A1", EndCell: "A5", Value: "Category"},
			},
			expected: []ParsedMergeRange{
				{StartRow: 0, StartCol: 0, EndRow: 4, EndCol: 0, Value: "Category"},
			},
		},
		{
			name: "single cell merge skipped",
			merges: []MergeCellInfo{
				{StartCell: "A1", EndCell: "A1", Value: "Single"},
			},
			expected: []ParsedMergeRange{},
		},
		{
			name: "multiple merges",
			merges: []MergeCellInfo{
				{StartCell: "A1", EndCell: "B1", Value: "Header1"},
				{StartCell: "C1", EndCell: "D1", Value: "Header2"},
			},
			expected: []ParsedMergeRange{
				{StartRow: 0, StartCol: 0, EndRow: 0, EndCol: 1, Value: "Header1"},
				{StartRow: 0, StartCol: 2, EndRow: 0, EndCol: 3, Value: "Header2"},
			},
		},
		{
			name:     "empty merges",
			merges:   []MergeCellInfo{},
			expected: []ParsedMergeRange{},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mp := NewMergeProcessor(models.DefaultConfig())
			result, err := mp.ParseMergeCells(tt.merges)

			if (err != nil) != tt.wantErr {
				t.Errorf("ParseMergeCells() error = %v, wantErr %v", err, tt.wantErr)
				return
			}

			if len(result) != len(tt.expected) {
				t.Errorf("ParseMergeCells() got %d results, want %d", len(result), len(tt.expected))
				return
			}

			for i, exp := range tt.expected {
				got := result[i]
				if got.StartRow != exp.StartRow || got.StartCol != exp.StartCol ||
					got.EndRow != exp.EndRow || got.EndCol != exp.EndCol ||
					got.Value != exp.Value {
					t.Errorf("ParseMergeCells()[%d] = %+v, want %+v", i, got, exp)
				}
			}
		})
	}
}

func TestMergeProcessor_ParseMergeCells_InvalidCell(t *testing.T) {
	mp := NewMergeProcessor(models.DefaultConfig())
	merges := []MergeCellInfo{
		{StartCell: "INVALID", EndCell: "B2", Value: "Test"},
	}

	_, err := mp.ParseMergeCells(merges)
	if err == nil {
		t.Error("Expected error for invalid cell reference")
	}
}

func TestMergeProcessor_ApplyMerges_Basic(t *testing.T) {
	config := models.DefaultConfig()
	mp := NewMergeProcessor(config)

	// Create a 3x3 grid
	grid := make([][]models.Cell, 3)
	for i := range grid {
		grid[i] = make([]models.Cell, 3)
		for j := range grid[i] {
			grid[i][j] = models.Cell{
				Row:      i,
				Col:      j,
				Value:    "",
				RawValue: "",
				Type:     models.CellTypeEmpty,
			}
		}
	}

	// Set origin cell value
	grid[0][0].Value = "Merged"
	grid[0][0].RawValue = "Merged"
	grid[0][0].Type = models.CellTypeString

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "Merged"},
	}

	mp.ApplyMerges(grid, merges)

	// Check all cells in the merge range
	for row := 0; row <= 1; row++ {
		for col := 0; col <= 1; col++ {
			cell := grid[row][col]

			if cell.RawValue != "Merged" {
				t.Errorf("Cell[%d][%d].RawValue = %q, want %q", row, col, cell.RawValue, "Merged")
			}
			if !cell.IsMerged {
				t.Errorf("Cell[%d][%d].IsMerged = false, want true", row, col)
			}
			if cell.MergeRange == nil {
				t.Errorf("Cell[%d][%d].MergeRange = nil, want non-nil", row, col)
				continue
			}
			if cell.MergeRange.StartRow != 0 || cell.MergeRange.StartCol != 0 ||
				cell.MergeRange.EndRow != 1 || cell.MergeRange.EndCol != 1 {
				t.Errorf("Cell[%d][%d].MergeRange incorrect", row, col)
			}

			isOrigin := row == 0 && col == 0
			if cell.MergeRange.IsOrigin != isOrigin {
				t.Errorf("Cell[%d][%d].MergeRange.IsOrigin = %v, want %v", row, col, cell.MergeRange.IsOrigin, isOrigin)
			}
		}
	}

	// Check cell outside merge is unchanged
	if grid[2][2].IsMerged {
		t.Error("Cell[2][2] should not be merged")
	}
}

func TestMergeProcessor_ApplyMerges_EmptyValue(t *testing.T) {
	config := models.DefaultConfig()
	mp := NewMergeProcessor(config)

	grid := make([][]models.Cell, 2)
	for i := range grid {
		grid[i] = make([]models.Cell, 2)
		for j := range grid[i] {
			grid[i][j] = models.Cell{Row: i, Col: j, Type: models.CellTypeEmpty}
		}
	}

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: ""},
	}

	mp.ApplyMerges(grid, merges)

	// All cells should be marked as merged even with empty value
	for row := 0; row <= 1; row++ {
		for col := 0; col <= 1; col++ {
			if !grid[row][col].IsMerged {
				t.Errorf("Cell[%d][%d] should be marked as merged even with empty value", row, col)
			}
		}
	}
}

func TestMergeProcessor_ApplyMerges_OutOfBounds(t *testing.T) {
	config := models.DefaultConfig()
	mp := NewMergeProcessor(config)

	// Small 2x2 grid
	grid := make([][]models.Cell, 2)
	for i := range grid {
		grid[i] = make([]models.Cell, 2)
		for j := range grid[i] {
			grid[i][j] = models.Cell{Row: i, Col: j, Type: models.CellTypeEmpty}
		}
	}

	// Merge extends beyond grid
	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 5, EndCol: 5, Value: "Big"},
	}

	// Should not panic
	mp.ApplyMerges(grid, merges)

	// Cells within grid should be merged
	for row := 0; row <= 1; row++ {
		for col := 0; col <= 1; col++ {
			if !grid[row][col].IsMerged {
				t.Errorf("Cell[%d][%d] should be merged", row, col)
			}
		}
	}
}

func TestMergeProcessor_ApplyMerges_ConfigDisabled(t *testing.T) {
	config := models.DetectionConfig{
		ExpandMergedCells:  false,
		TrackMergeMetadata: false,
	}
	mp := NewMergeProcessor(config)

	grid := make([][]models.Cell, 2)
	for i := range grid {
		grid[i] = make([]models.Cell, 2)
		for j := range grid[i] {
			grid[i][j] = models.Cell{
				Row:      i,
				Col:      j,
				Value:    "Original",
				RawValue: "Original",
				Type:     models.CellTypeString,
			}
		}
	}

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "Merged"},
	}

	mp.ApplyMerges(grid, merges)

	// With both options disabled, cells should be unchanged
	for row := 0; row <= 1; row++ {
		for col := 0; col <= 1; col++ {
			cell := grid[row][col]
			if cell.IsMerged {
				t.Errorf("Cell[%d][%d].IsMerged should be false when TrackMergeMetadata is disabled", row, col)
			}
			if cell.RawValue != "Original" {
				t.Errorf("Cell[%d][%d].RawValue should be unchanged when ExpandMergedCells is disabled", row, col)
			}
		}
	}
}

func TestMergeProcessor_ApplyMerges_OnlyTracking(t *testing.T) {
	config := models.DetectionConfig{
		ExpandMergedCells:  false,
		TrackMergeMetadata: true,
	}
	mp := NewMergeProcessor(config)

	grid := make([][]models.Cell, 2)
	for i := range grid {
		grid[i] = make([]models.Cell, 2)
		for j := range grid[i] {
			grid[i][j] = models.Cell{
				Row:      i,
				Col:      j,
				Value:    "Original",
				RawValue: "Original",
				Type:     models.CellTypeString,
			}
		}
	}

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "Merged"},
	}

	mp.ApplyMerges(grid, merges)

	// Metadata should be set, but values unchanged
	for row := 0; row <= 1; row++ {
		for col := 0; col <= 1; col++ {
			cell := grid[row][col]
			if !cell.IsMerged {
				t.Errorf("Cell[%d][%d].IsMerged should be true when TrackMergeMetadata is enabled", row, col)
			}
			if cell.RawValue != "Original" {
				t.Errorf("Cell[%d][%d].RawValue should be unchanged when ExpandMergedCells is disabled", row, col)
			}
		}
	}
}

func TestMergeProcessor_ApplyMerges_EmptyGrid(t *testing.T) {
	config := models.DefaultConfig()
	mp := NewMergeProcessor(config)

	var grid [][]models.Cell // nil grid

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "Merged"},
	}

	// Should not panic
	mp.ApplyMerges(grid, merges)
}

func TestMergeProcessor_BuildMergeMap(t *testing.T) {
	mp := NewMergeProcessor(models.DefaultConfig())

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "A"},
		{StartRow: 0, StartCol: 3, EndRow: 0, EndCol: 4, Value: "B"},
	}

	mergeMap := mp.BuildMergeMap(merges)

	// Check cells in first merge
	for row := 0; row <= 1; row++ {
		for col := 0; col <= 1; col++ {
			if mergeMap[row][col] == nil {
				t.Errorf("mergeMap[%d][%d] should not be nil", row, col)
			} else if mergeMap[row][col].Value != "A" {
				t.Errorf("mergeMap[%d][%d].Value = %q, want %q", row, col, mergeMap[row][col].Value, "A")
			}
		}
	}

	// Check cells in second merge
	if mergeMap[0][3] == nil || mergeMap[0][3].Value != "B" {
		t.Error("mergeMap[0][3] should point to merge B")
	}
	if mergeMap[0][4] == nil || mergeMap[0][4].Value != "B" {
		t.Error("mergeMap[0][4] should point to merge B")
	}

	// Check cell not in any merge
	if mergeMap[0][2] != nil {
		t.Error("mergeMap[0][2] should be nil (not in any merge)")
	}
}

func TestMergeProcessor_IsCellInMerge(t *testing.T) {
	mp := NewMergeProcessor(models.DefaultConfig())

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "A"},
		{StartRow: 5, StartCol: 5, EndRow: 7, EndCol: 7, Value: "B"},
	}

	tests := []struct {
		row, col int
		inMerge  bool
		value    string
	}{
		{0, 0, true, "A"},
		{1, 1, true, "A"},
		{0, 1, true, "A"},
		{2, 2, false, ""},
		{5, 5, true, "B"},
		{6, 6, true, "B"},
		{4, 4, false, ""},
	}

	for _, tt := range tests {
		result := mp.IsCellInMerge(merges, tt.row, tt.col)
		if tt.inMerge {
			if result == nil {
				t.Errorf("IsCellInMerge(%d, %d) = nil, want merge with value %q", tt.row, tt.col, tt.value)
			} else if result.Value != tt.value {
				t.Errorf("IsCellInMerge(%d, %d).Value = %q, want %q", tt.row, tt.col, result.Value, tt.value)
			}
		} else {
			if result != nil {
				t.Errorf("IsCellInMerge(%d, %d) = %+v, want nil", tt.row, tt.col, result)
			}
		}
	}
}

func TestMergeProcessor_ApplyMerges_PreservesType(t *testing.T) {
	config := models.DefaultConfig()
	mp := NewMergeProcessor(config)

	grid := make([][]models.Cell, 2)
	for i := range grid {
		grid[i] = make([]models.Cell, 2)
		for j := range grid[i] {
			grid[i][j] = models.Cell{Row: i, Col: j, Type: models.CellTypeEmpty}
		}
	}

	// Set origin with number type
	grid[0][0].Value = float64(42)
	grid[0][0].RawValue = "42"
	grid[0][0].Type = models.CellTypeNumber

	merges := []ParsedMergeRange{
		{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 1, Value: "42"},
	}

	mp.ApplyMerges(grid, merges)

	// Non-origin cells should inherit the type from origin
	for row := 0; row <= 1; row++ {
		for col := 0; col <= 1; col++ {
			cell := grid[row][col]
			if cell.Type != models.CellTypeNumber {
				t.Errorf("Cell[%d][%d].Type = %v, want CellTypeNumber", row, col, cell.Type)
			}
		}
	}
}
