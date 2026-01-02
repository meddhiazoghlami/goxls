package reader

import (
	"strings"
	"testing"

	"excel-lite/pkg/models"
)

// =============================================================================
// Constructor Tests
// =============================================================================

func TestNewHeaderDetector(t *testing.T) {
	config := models.DetectionConfig{
		MinColumns:    4,
		HeaderDensity: 0.8,
	}

	hd := NewHeaderDetector(config)

	if hd == nil {
		t.Fatal("NewHeaderDetector() returned nil")
	}

	if hd.config.MinColumns != 4 {
		t.Errorf("config.MinColumns = %d, want 4", hd.config.MinColumns)
	}

	if hd.config.HeaderDensity != 0.8 {
		t.Errorf("config.HeaderDensity = %v, want 0.8", hd.config.HeaderDensity)
	}
}

func TestNewDefaultHeaderDetector(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	if hd == nil {
		t.Fatal("NewDefaultHeaderDetector() returned nil")
	}

	defaultConfig := models.DefaultConfig()
	if hd.config.MinColumns != defaultConfig.MinColumns {
		t.Errorf("config.MinColumns = %d, want %d", hd.config.MinColumns, defaultConfig.MinColumns)
	}
}

// =============================================================================
// DetectHeaderRow Tests
// =============================================================================

func TestHeaderDetector_DetectHeaderRow_Simple(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Grid with headers in row 0
	grid := [][]models.Cell{
		{
			makeCell("Name", models.CellTypeString),
			makeCell("Age", models.CellTypeString),
			makeCell("Email", models.CellTypeString),
		},
		{
			makeCell("Alice", models.CellTypeString),
			makeCell("30", models.CellTypeNumber),
			makeCell("alice@test.com", models.CellTypeString),
		},
		{
			makeCell("Bob", models.CellTypeString),
			makeCell("25", models.CellTypeNumber),
			makeCell("bob@test.com", models.CellTypeString),
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2}

	headerRow := hd.DetectHeaderRow(grid, boundary)

	if headerRow != 0 {
		t.Errorf("DetectHeaderRow() = %d, want 0", headerRow)
	}
}

func TestHeaderDetector_DetectHeaderRow_OffsetHeader(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Grid with metadata in rows 0-1, headers in row 2
	grid := [][]models.Cell{
		{makeCell("Report", models.CellTypeString), makeEmptyCell(), makeEmptyCell()},
		{makeEmptyCell(), makeEmptyCell(), makeEmptyCell()},
		{makeCell("ID", models.CellTypeString), makeCell("Name", models.CellTypeString), makeCell("Status", models.CellTypeString)},
		{makeCell("1", models.CellTypeNumber), makeCell("Alice", models.CellTypeString), makeCell("Active", models.CellTypeString)},
		{makeCell("2", models.CellTypeNumber), makeCell("Bob", models.CellTypeString), makeCell("Inactive", models.CellTypeString)},
	}

	boundary := models.TableBoundary{StartRow: 2, EndRow: 4, StartCol: 0, EndCol: 2}

	headerRow := hd.DetectHeaderRow(grid, boundary)

	if headerRow != 2 {
		t.Errorf("DetectHeaderRow() = %d, want 2", headerRow)
	}
}

func TestHeaderDetector_DetectHeaderRow_CommonPatterns(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Grid with common header names
	grid := [][]models.Cell{
		{
			makeCell("First Name", models.CellTypeString),
			makeCell("Last Name", models.CellTypeString),
			makeCell("Email Address", models.CellTypeString),
			makeCell("Phone Number", models.CellTypeString),
		},
		{
			makeCell("John", models.CellTypeString),
			makeCell("Doe", models.CellTypeString),
			makeCell("john@example.com", models.CellTypeString),
			makeCell("555-1234", models.CellTypeString),
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 3}

	headerRow := hd.DetectHeaderRow(grid, boundary)

	if headerRow != 0 {
		t.Errorf("DetectHeaderRow() = %d, want 0", headerRow)
	}
}

func TestHeaderDetector_DetectHeaderRow_NumbersVsStrings(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Header row with strings, data rows with numbers
	grid := [][]models.Cell{
		{makeCell("A", models.CellTypeString), makeCell("B", models.CellTypeString), makeCell("C", models.CellTypeString)},
		{makeCell("1", models.CellTypeNumber), makeCell("2", models.CellTypeNumber), makeCell("3", models.CellTypeNumber)},
		{makeCell("4", models.CellTypeNumber), makeCell("5", models.CellTypeNumber), makeCell("6", models.CellTypeNumber)},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2}

	headerRow := hd.DetectHeaderRow(grid, boundary)

	if headerRow != 0 {
		t.Errorf("DetectHeaderRow() = %d, want 0 (string row)", headerRow)
	}
}

func TestHeaderDetector_DetectHeaderRow_OutOfBounds(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("A", models.CellTypeString)},
	}

	// Boundary beyond grid
	boundary := models.TableBoundary{StartRow: 10, EndRow: 15, StartCol: 0, EndCol: 0}

	// Should return the start row (even if out of bounds)
	headerRow := hd.DetectHeaderRow(grid, boundary)

	if headerRow != 10 {
		t.Errorf("DetectHeaderRow() = %d, want 10", headerRow)
	}
}

// =============================================================================
// ExtractHeaders Tests
// =============================================================================

func TestHeaderDetector_ExtractHeaders_Simple(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{
			makeCell("Name", models.CellTypeString),
			makeCell("Age", models.CellTypeString),
			makeCell("City", models.CellTypeString),
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2}

	headers := hd.ExtractHeaders(grid, 0, boundary)

	expected := []string{"Name", "Age", "City"}
	if len(headers) != len(expected) {
		t.Fatalf("ExtractHeaders() returned %d headers, want %d", len(headers), len(expected))
	}

	for i, exp := range expected {
		if headers[i] != exp {
			t.Errorf("headers[%d] = %q, want %q", i, headers[i], exp)
		}
	}
}

func TestHeaderDetector_ExtractHeaders_EmptyHeaders(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{
			makeCell("Name", models.CellTypeString),
			makeEmptyCell(),
			makeCell("Age", models.CellTypeString),
			makeEmptyCell(),
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 3}

	headers := hd.ExtractHeaders(grid, 0, boundary)

	if headers[0] != "Name" {
		t.Errorf("headers[0] = %q, want Name", headers[0])
	}

	// Empty cells should get generated names
	if !strings.HasPrefix(headers[1], "Column_") {
		t.Errorf("headers[1] = %q, expected Column_* for empty header", headers[1])
	}

	if headers[2] != "Age" {
		t.Errorf("headers[2] = %q, want Age", headers[2])
	}
}

func TestHeaderDetector_ExtractHeaders_Duplicates(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{
			makeCell("Name", models.CellTypeString),
			makeCell("Name", models.CellTypeString),
			makeCell("Name", models.CellTypeString),
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2}

	headers := hd.ExtractHeaders(grid, 0, boundary)

	// Should have unique names
	seen := make(map[string]bool)
	for _, h := range headers {
		if seen[h] {
			t.Errorf("Duplicate header found: %q", h)
		}
		seen[h] = true
	}

	// First one should be original
	if headers[0] != "Name" {
		t.Errorf("headers[0] = %q, want Name", headers[0])
	}

	// Others should be suffixed
	if !strings.HasPrefix(headers[1], "Name_") {
		t.Errorf("headers[1] = %q, expected Name_* suffix", headers[1])
	}
}

func TestHeaderDetector_ExtractHeaders_Whitespace(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{
			makeCell("  Name  ", models.CellTypeString),
			makeCell("\tAge\t", models.CellTypeString),
			makeCell("City", models.CellTypeString),
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2}

	headers := hd.ExtractHeaders(grid, 0, boundary)

	// Should be trimmed
	if headers[0] != "Name" {
		t.Errorf("headers[0] = %q, want Name (trimmed)", headers[0])
	}

	if headers[1] != "Age" {
		t.Errorf("headers[1] = %q, want Age (trimmed)", headers[1])
	}
}

func TestHeaderDetector_ExtractHeaders_OutOfBounds(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("A", models.CellTypeString)},
	}

	headers := hd.ExtractHeaders(grid, 10, models.TableBoundary{})

	if headers != nil {
		t.Errorf("ExtractHeaders() with out-of-bounds row should return nil, got %v", headers)
	}
}

func TestHeaderDetector_ExtractHeaders_PartialBoundary(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{
			makeCell("A", models.CellTypeString),
			makeCell("B", models.CellTypeString),
			makeCell("C", models.CellTypeString),
			makeCell("D", models.CellTypeString),
		},
	}

	// Only extract columns 1-2
	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 1, EndCol: 2}

	headers := hd.ExtractHeaders(grid, 0, boundary)

	if len(headers) != 2 {
		t.Fatalf("ExtractHeaders() returned %d headers, want 2", len(headers))
	}

	if headers[0] != "B" || headers[1] != "C" {
		t.Errorf("headers = %v, want [B C]", headers)
	}
}

// =============================================================================
// normalizeHeader Tests
// =============================================================================

func TestHeaderDetector_normalizeHeader(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	tests := []struct {
		name      string
		value     string
		colIndex  int
		usedNames map[string]int
		expected  string
	}{
		{
			name:      "simple",
			value:     "Name",
			colIndex:  0,
			usedNames: map[string]int{},
			expected:  "Name",
		},
		{
			name:      "with whitespace",
			value:     "  First Name  ",
			colIndex:  0,
			usedNames: map[string]int{},
			expected:  "First Name",
		},
		{
			name:      "empty generates column name",
			value:     "",
			colIndex:  2,
			usedNames: map[string]int{},
			expected:  "Column_3",
		},
		{
			name:      "duplicate gets suffix",
			value:     "Name",
			colIndex:  1,
			usedNames: map[string]int{"name": 1},
			expected:  "Name_2",
		},
		{
			name:      "case insensitive duplicate",
			value:     "NAME",
			colIndex:  1,
			usedNames: map[string]int{"name": 1},
			expected:  "NAME_2",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := hd.normalizeHeader(tt.value, tt.colIndex, tt.usedNames)
			if got != tt.expected {
				t.Errorf("normalizeHeader(%q) = %q, want %q", tt.value, got, tt.expected)
			}
		})
	}
}

// =============================================================================
// ValidateHeaders Tests
// =============================================================================

func TestHeaderDetector_ValidateHeaders(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	tests := []struct {
		name     string
		headers  []string
		expected bool
	}{
		{
			name:     "valid headers",
			headers:  []string{"Name", "Age", "Email"},
			expected: true,
		},
		{
			name:     "minimum valid",
			headers:  []string{"A", "B"},
			expected: true,
		},
		{
			name:     "too few columns",
			headers:  []string{"Only"},
			expected: false,
		},
		{
			name:     "empty",
			headers:  []string{},
			expected: false,
		},
		{
			name:     "mostly generated",
			headers:  []string{"Column_1", "Column_2", "Column_3", "Column_4"},
			expected: false,
		},
		{
			name:     "half meaningful",
			headers:  []string{"Name", "Column_2", "Age", "Column_4"},
			expected: true,
		},
		{
			name:     "just under threshold",
			headers:  []string{"Name", "Column_2", "Column_3", "Column_4"},
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := hd.ValidateHeaders(tt.headers); got != tt.expected {
				t.Errorf("ValidateHeaders(%v) = %v, want %v", tt.headers, got, tt.expected)
			}
		})
	}
}

// =============================================================================
// scoreAsHeader Tests
// =============================================================================

func TestHeaderDetector_scoreAsHeader(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	tests := []struct {
		name         string
		grid         [][]models.Cell
		row          int
		boundary     models.TableBoundary
		expectHigher bool // Should score higher than a data row
	}{
		{
			name: "all strings header",
			grid: [][]models.Cell{
				{makeCell("Name", models.CellTypeString), makeCell("Age", models.CellTypeString)},
				{makeCell("Alice", models.CellTypeString), makeCell("30", models.CellTypeNumber)},
			},
			row:          0,
			boundary:     models.TableBoundary{StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 1},
			expectHigher: true,
		},
		{
			name: "common header words",
			grid: [][]models.Cell{
				{makeCell("ID", models.CellTypeString), makeCell("Name", models.CellTypeString), makeCell("Email", models.CellTypeString)},
				{makeCell("1", models.CellTypeNumber), makeCell("Test", models.CellTypeString), makeCell("test@test.com", models.CellTypeString)},
			},
			row:          0,
			boundary:     models.TableBoundary{StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 2},
			expectHigher: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			headerScore := hd.scoreAsHeader(tt.grid, tt.row, tt.boundary)
			dataScore := hd.scoreAsHeader(tt.grid, tt.row+1, tt.boundary)

			if tt.expectHigher && headerScore <= dataScore {
				t.Errorf("Header row score (%v) should be higher than data row score (%v)", headerScore, dataScore)
			}
		})
	}
}

// =============================================================================
// countStrings Tests
// =============================================================================

func TestHeaderDetector_countStrings(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	tests := []struct {
		name     string
		grid     [][]models.Cell
		row      int
		boundary models.TableBoundary
		expected int
	}{
		{
			name: "all strings",
			grid: [][]models.Cell{
				{makeCell("A", models.CellTypeString), makeCell("B", models.CellTypeString), makeCell("C", models.CellTypeString)},
			},
			row:      0,
			boundary: models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2},
			expected: 3,
		},
		{
			name: "mixed types",
			grid: [][]models.Cell{
				{makeCell("A", models.CellTypeString), makeCell("1", models.CellTypeNumber), makeCell("true", models.CellTypeBool)},
			},
			row:      0,
			boundary: models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2},
			expected: 1,
		},
		{
			name: "no strings",
			grid: [][]models.Cell{
				{makeCell("1", models.CellTypeNumber), makeCell("2", models.CellTypeNumber)},
			},
			row:      0,
			boundary: models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 1},
			expected: 0,
		},
		{
			name:     "out of bounds",
			grid:     [][]models.Cell{},
			row:      5,
			boundary: models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2},
			expected: 0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := hd.countStrings(tt.grid, tt.row, tt.boundary); got != tt.expected {
				t.Errorf("countStrings() = %d, want %d", got, tt.expected)
			}
		})
	}
}

// =============================================================================
// patternBonus Tests
// =============================================================================

func TestHeaderDetector_patternBonus(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	tests := []struct {
		name        string
		grid        [][]models.Cell
		row         int
		boundary    models.TableBoundary
		expectBonus bool
	}{
		{
			name: "common patterns",
			grid: [][]models.Cell{
				{makeCell("ID", models.CellTypeString), makeCell("Name", models.CellTypeString), makeCell("Email", models.CellTypeString)},
			},
			row:         0,
			boundary:    models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2},
			expectBonus: true,
		},
		{
			name: "no common patterns",
			grid: [][]models.Cell{
				{makeCell("XYZ", models.CellTypeString), makeCell("ABC", models.CellTypeString), makeCell("123", models.CellTypeString)},
			},
			row:         0,
			boundary:    models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2},
			expectBonus: false,
		},
		{
			name: "partial patterns",
			grid: [][]models.Cell{
				{makeCell("First Name", models.CellTypeString), makeCell("Random", models.CellTypeString), makeCell("Created Date", models.CellTypeString)},
			},
			row:         0,
			boundary:    models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2},
			expectBonus: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			bonus := hd.patternBonus(tt.grid, tt.row, tt.boundary)
			hasBonus := bonus > 0
			if hasBonus != tt.expectBonus {
				t.Errorf("patternBonus() returned %v, expectBonus = %v", bonus, tt.expectBonus)
			}
		})
	}
}

func TestHeaderDetector_patternBonus_OutOfBounds(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("Name", models.CellTypeString)},
	}

	// Row out of bounds
	bonus := hd.patternBonus(grid, 10, models.TableBoundary{})
	if bonus != 0 {
		t.Errorf("patternBonus() with out-of-bounds row should return 0, got %v", bonus)
	}
}

func TestHeaderDetector_patternBonus_EmptyCells(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeEmptyCell(), makeEmptyCell(), makeEmptyCell()},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 2}
	bonus := hd.patternBonus(grid, 0, boundary)

	// Should return 0 for empty cells
	if bonus != 0 {
		t.Errorf("patternBonus() with empty cells should return 0, got %v", bonus)
	}
}

func TestHeaderDetector_scoreAsHeader_OutOfBounds(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("A", models.CellTypeString)},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 0}

	// Row out of bounds
	score := hd.scoreAsHeader(grid, 10, boundary)
	if score != 0 {
		t.Errorf("scoreAsHeader() with out-of-bounds row should return 0, got %v", score)
	}
}

func TestHeaderDetector_scoreAsHeader_EmptyRow(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeEmptyCell(), makeEmptyCell()},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 1}

	score := hd.scoreAsHeader(grid, 0, boundary)
	// Empty rows should have low or zero score
	if score > 50 {
		t.Errorf("scoreAsHeader() for empty row should have low score, got %v", score)
	}
}

func TestHeaderDetector_scoreAsHeader_NoNextRow(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Single row - no next row to compare
	grid := [][]models.Cell{
		{makeCell("Name", models.CellTypeString), makeCell("Age", models.CellTypeString)},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 1}

	score := hd.scoreAsHeader(grid, 0, boundary)
	// Should still calculate a score without the next row comparison
	if score <= 0 {
		t.Errorf("scoreAsHeader() should return positive score even without next row, got %v", score)
	}
}

func TestHeaderDetector_scoreAsHeader_WithMergedCells(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Grid with merged cells in header
	grid := [][]models.Cell{
		{
			{Value: "Group", Type: models.CellTypeString, RawValue: "Group", IsMerged: true, MergeRange: &models.MergeRange{StartRow: 0, StartCol: 0, EndRow: 0, EndCol: 1, IsOrigin: true}},
			{Value: "Group", Type: models.CellTypeString, RawValue: "Group", IsMerged: true, MergeRange: &models.MergeRange{StartRow: 0, StartCol: 0, EndRow: 0, EndCol: 1, IsOrigin: false}},
			{Value: "Other", Type: models.CellTypeString, RawValue: "Other"},
		},
		{
			{Value: "Data1", Type: models.CellTypeString, RawValue: "Data1"},
			{Value: "Data2", Type: models.CellTypeString, RawValue: "Data2"},
			{Value: "Data3", Type: models.CellTypeString, RawValue: "Data3"},
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 2}

	score := hd.scoreAsHeader(grid, 0, boundary)
	// Should get bonus for merged cells
	if score <= 0 {
		t.Errorf("scoreAsHeader() with merged cells should return positive score, got %v", score)
	}
}

// =============================================================================
// DetectHeaderRows Tests
// =============================================================================

func TestHeaderDetector_DetectHeaderRows_SingleRow(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("Name", models.CellTypeString), makeCell("Age", models.CellTypeString)},
		{makeCell("Alice", models.CellTypeString), makeCell("30", models.CellTypeNumber)},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 1, StartCol: 0, EndCol: 1}

	start, end := hd.DetectHeaderRows(grid, boundary)
	if start != 0 {
		t.Errorf("DetectHeaderRows() start = %d, want 0", start)
	}
	if end != 0 {
		t.Errorf("DetectHeaderRows() end = %d, want 0", end)
	}
}

func TestHeaderDetector_DetectHeaderRows_WithMergedMultiRow(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Grid with merged cell spanning multiple rows
	grid := [][]models.Cell{
		{
			{Value: "Category", Type: models.CellTypeString, RawValue: "Category", IsMerged: true, MergeRange: &models.MergeRange{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 0, IsOrigin: true}},
			{Value: "Sub1", Type: models.CellTypeString, RawValue: "Sub1"},
		},
		{
			{Value: "Category", Type: models.CellTypeString, RawValue: "Category", IsMerged: true, MergeRange: &models.MergeRange{StartRow: 0, StartCol: 0, EndRow: 1, EndCol: 0, IsOrigin: false}},
			{Value: "Sub2", Type: models.CellTypeString, RawValue: "Sub2"},
		},
		{
			{Value: "Data1", Type: models.CellTypeString, RawValue: "Data1"},
			{Value: "Data2", Type: models.CellTypeString, RawValue: "Data2"},
		},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 1}

	start, end := hd.DetectHeaderRows(grid, boundary)
	if start != 0 {
		t.Errorf("DetectHeaderRows() start = %d, want 0", start)
	}
	if end != 1 {
		t.Errorf("DetectHeaderRows() end = %d, want 1", end)
	}
}

func TestHeaderDetector_DetectHeaderRows_EmptyGrid(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 0}

	start, end := hd.DetectHeaderRows(grid, boundary)
	if start != 0 || end != 0 {
		t.Errorf("DetectHeaderRows() on empty grid should return (0, 0), got (%d, %d)", start, end)
	}
}

func TestHeaderDetector_DetectHeaderRows_LimitMaxRows(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Create grid with merge spanning 5 rows (exceeds max of 3)
	grid := make([][]models.Cell, 6)
	for i := 0; i < 6; i++ {
		grid[i] = []models.Cell{
			{Value: "Header", Type: models.CellTypeString, RawValue: "Header", IsMerged: true, MergeRange: &models.MergeRange{StartRow: 0, StartCol: 0, EndRow: 4, EndCol: 0, IsOrigin: i == 0}},
		}
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 5, StartCol: 0, EndCol: 0}

	start, end := hd.DetectHeaderRows(grid, boundary)
	// Should be limited to max 3 rows (0, 1, 2)
	if end-start >= 3 {
		t.Errorf("DetectHeaderRows() should limit to 3 rows max, got %d rows", end-start+1)
	}
}

// =============================================================================
// ExtractHierarchicalHeaders Tests
// =============================================================================

func TestHeaderDetector_ExtractHierarchicalHeaders(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("Group A", models.CellTypeString), makeCell("Group A", models.CellTypeString), makeCell("Group B", models.CellTypeString)},
		{makeCell("Sub1", models.CellTypeString), makeCell("Sub2", models.CellTypeString), makeCell("Sub3", models.CellTypeString)},
		{makeCell("Data", models.CellTypeString), makeCell("Data", models.CellTypeString), makeCell("Data", models.CellTypeString)},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 2}

	result := hd.ExtractHierarchicalHeaders(grid, 0, 1, boundary)

	if len(result) != 2 {
		t.Fatalf("ExtractHierarchicalHeaders() returned %d levels, want 2", len(result))
	}

	if len(result[0]) != 3 {
		t.Errorf("Level 0 has %d columns, want 3", len(result[0]))
	}

	if len(result[1]) != 3 {
		t.Errorf("Level 1 has %d columns, want 3", len(result[1]))
	}

	if result[0][0] != "Group A" {
		t.Errorf("result[0][0] = %q, want %q", result[0][0], "Group A")
	}

	if result[1][0] != "Sub1" {
		t.Errorf("result[1][0] = %q, want %q", result[1][0], "Sub1")
	}
}

func TestHeaderDetector_ExtractHierarchicalHeaders_InvalidRange(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("Header", models.CellTypeString)},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 0}

	// Invalid range where headerEnd < headerStart
	result := hd.ExtractHierarchicalHeaders(grid, 5, 3, boundary)

	if result != nil {
		t.Errorf("ExtractHierarchicalHeaders() with invalid range should return nil, got %v", result)
	}
}

func TestHeaderDetector_ExtractHierarchicalHeaders_OutOfBoundsRow(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	grid := [][]models.Cell{
		{makeCell("Header", models.CellTypeString)},
	}

	boundary := models.TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 0}

	// headerStart is within grid but headerEnd is beyond
	result := hd.ExtractHierarchicalHeaders(grid, 0, 5, boundary)

	if len(result) != 6 {
		t.Errorf("ExtractHierarchicalHeaders() returned %d levels, want 6", len(result))
	}
}

// =============================================================================
// FlattenHierarchicalHeaders Tests
// =============================================================================

func TestHeaderDetector_FlattenHierarchicalHeaders(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	hierarchical := [][]string{
		{"Group A", "Group A", "Group B"},
		{"Sub1", "Sub2", "Sub3"},
	}

	result := hd.FlattenHierarchicalHeaders(hierarchical, " > ")

	if len(result) != 3 {
		t.Fatalf("FlattenHierarchicalHeaders() returned %d columns, want 3", len(result))
	}

	expected := []string{"Group A > Sub1", "Group A > Sub2", "Group B > Sub3"}
	for i, exp := range expected {
		if result[i] != exp {
			t.Errorf("result[%d] = %q, want %q", i, result[i], exp)
		}
	}
}

func TestHeaderDetector_FlattenHierarchicalHeaders_NilInput(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	result := hd.FlattenHierarchicalHeaders(nil, " > ")

	if result != nil {
		t.Errorf("FlattenHierarchicalHeaders(nil) should return nil, got %v", result)
	}
}

func TestHeaderDetector_FlattenHierarchicalHeaders_EmptyInput(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	result := hd.FlattenHierarchicalHeaders([][]string{}, " > ")

	if result != nil {
		t.Errorf("FlattenHierarchicalHeaders([]) should return nil, got %v", result)
	}
}

func TestHeaderDetector_FlattenHierarchicalHeaders_DuplicateRemoval(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Same header appears in both levels (common with merged cells)
	hierarchical := [][]string{
		{"Category", "Category", "Other"},
		{"Category", "SubCat", "Other"},
	}

	result := hd.FlattenHierarchicalHeaders(hierarchical, " > ")

	// "Category" should not be duplicated
	if result[0] != "Category" {
		t.Errorf("result[0] = %q, want %q (duplicates should be removed)", result[0], "Category")
	}
	if result[1] != "Category > SubCat" {
		t.Errorf("result[1] = %q, want %q", result[1], "Category > SubCat")
	}
}

func TestHeaderDetector_FlattenHierarchicalHeaders_GenericColumnNames(t *testing.T) {
	hd := NewDefaultHeaderDetector()

	// Empty headers result in Column_N names
	hierarchical := [][]string{
		{"", "", ""},
	}

	result := hd.FlattenHierarchicalHeaders(hierarchical, " > ")

	for i, r := range result {
		if !strings.HasPrefix(r, "Column_") {
			t.Errorf("result[%d] = %q, should start with Column_", i, r)
		}
	}
}
