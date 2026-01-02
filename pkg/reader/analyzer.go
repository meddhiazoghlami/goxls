package reader

import (
	"excel-lite/pkg/models"
)

// TableAnalyzer detects table boundaries within a sheet
type TableAnalyzer struct {
	config models.DetectionConfig
}

// NewTableAnalyzer creates a new table analyzer with the given config
func NewTableAnalyzer(config models.DetectionConfig) *TableAnalyzer {
	return &TableAnalyzer{config: config}
}

// NewDefaultAnalyzer creates a table analyzer with default configuration
func NewDefaultAnalyzer() *TableAnalyzer {
	return NewTableAnalyzer(models.DefaultConfig())
}

// DetectTables finds all tables within a cell grid
func (ta *TableAnalyzer) DetectTables(grid [][]models.Cell) []models.TableBoundary {
	if len(grid) == 0 {
		return nil
	}

	var tables []models.TableBoundary
	visited := make([][]bool, len(grid))
	for i := range visited {
		if len(grid[i]) > 0 {
			visited[i] = make([]bool, len(grid[i]))
		}
	}

	// Scan for table start points
	for rowIdx := 0; rowIdx < len(grid); rowIdx++ {
		for colIdx := 0; colIdx < len(grid[rowIdx]); colIdx++ {
			if visited[rowIdx][colIdx] {
				continue
			}

			if !grid[rowIdx][colIdx].IsEmpty() {
				// Found a potential table start
				boundary := ta.expandTable(grid, rowIdx, colIdx, visited)
				if ta.isValidTable(boundary) {
					tables = append(tables, boundary)
				}
			}
		}
	}

	return tables
}

// expandTable expands from a starting cell to find the full table boundary
func (ta *TableAnalyzer) expandTable(grid [][]models.Cell, startRow, startCol int, visited [][]bool) models.TableBoundary {
	maxRows := len(grid)
	maxCols := 0
	for _, row := range grid {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	// Find the left boundary (first non-empty column in contiguous block)
	leftCol := startCol
	for col := startCol - 1; col >= 0; col-- {
		if startRow < len(grid) && col < len(grid[startRow]) && !grid[startRow][col].IsEmpty() {
			leftCol = col
		} else {
			break
		}
	}

	// Find the right boundary
	rightCol := startCol
	consecutiveEmpty := 0
	for col := startCol + 1; col < maxCols; col++ {
		hasData := false
		// Check multiple rows to determine if column has data
		for row := startRow; row < minInt(startRow+10, maxRows); row++ {
			if row < len(grid) && col < len(grid[row]) && !grid[row][col].IsEmpty() {
				hasData = true
				break
			}
		}
		if hasData {
			rightCol = col
			consecutiveEmpty = 0
		} else {
			consecutiveEmpty++
			if consecutiveEmpty > 1 {
				break
			}
		}
	}

	// Find the bottom boundary
	endRow := startRow
	emptyRowCount := 0
	for row := startRow + 1; row < maxRows; row++ {
		rowHasData := false
		for col := leftCol; col <= rightCol && col < len(grid[row]); col++ {
			if !grid[row][col].IsEmpty() {
				rowHasData = true
				break
			}
		}
		if rowHasData {
			endRow = row
			emptyRowCount = 0
		} else {
			emptyRowCount++
			if emptyRowCount > ta.config.MaxEmptyRows {
				break
			}
		}
	}

	boundary := models.TableBoundary{
		StartRow: startRow,
		EndRow:   endRow,
		StartCol: leftCol,
		EndCol:   rightCol,
	}

	// Expand boundary to include full merge ranges
	boundary = ta.expandForMerges(grid, boundary)

	// Mark cells as visited (including expanded merge areas)
	for row := boundary.StartRow; row <= boundary.EndRow; row++ {
		for col := boundary.StartCol; col <= boundary.EndCol; col++ {
			if row < len(visited) && col < len(visited[row]) {
				visited[row][col] = true
			}
		}
	}

	return boundary
}

// expandForMerges ensures merged cells don't get cut off at table boundaries
func (ta *TableAnalyzer) expandForMerges(grid [][]models.Cell, boundary models.TableBoundary) models.TableBoundary {
	expanded := boundary

	// Check all cells at the boundary for merges that extend beyond
	for row := boundary.StartRow; row <= boundary.EndRow && row < len(grid); row++ {
		for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[row]); col++ {
			cell := grid[row][col]
			if cell.MergeRange != nil {
				// Expand boundary to include full merge
				if cell.MergeRange.StartRow < expanded.StartRow {
					expanded.StartRow = cell.MergeRange.StartRow
				}
				if cell.MergeRange.EndRow > expanded.EndRow {
					expanded.EndRow = cell.MergeRange.EndRow
				}
				if cell.MergeRange.StartCol < expanded.StartCol {
					expanded.StartCol = cell.MergeRange.StartCol
				}
				if cell.MergeRange.EndCol > expanded.EndCol {
					expanded.EndCol = cell.MergeRange.EndCol
				}
			}
		}
	}

	return expanded
}

// isValidTable checks if a boundary represents a valid table
func (ta *TableAnalyzer) isValidTable(boundary models.TableBoundary) bool {
	rowCount := boundary.EndRow - boundary.StartRow + 1
	colCount := boundary.EndCol - boundary.StartCol + 1

	return rowCount >= ta.config.MinRows && colCount >= ta.config.MinColumns
}

// FindDenseRegions finds regions with high cell density (likely tables)
func (ta *TableAnalyzer) FindDenseRegions(grid [][]models.Cell, windowSize int) []models.TableBoundary {
	if len(grid) == 0 || windowSize < 1 {
		return nil
	}

	var regions []models.TableBoundary
	maxRows := len(grid)
	maxCols := 0
	for _, row := range grid {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	// Sliding window to find dense regions
	for startRow := 0; startRow <= maxRows-windowSize; startRow++ {
		for startCol := 0; startCol <= maxCols-windowSize; startCol++ {
			density := ta.calculateDensity(grid, startRow, startCol, windowSize)
			if density >= ta.config.HeaderDensity {
				regions = append(regions, models.TableBoundary{
					StartRow: startRow,
					EndRow:   startRow + windowSize - 1,
					StartCol: startCol,
					EndCol:   startCol + windowSize - 1,
				})
			}
		}
	}

	return ta.mergeOverlappingRegions(regions)
}

// calculateDensity calculates the density of non-empty cells in a region
func (ta *TableAnalyzer) calculateDensity(grid [][]models.Cell, startRow, startCol, size int) float64 {
	total := 0
	nonEmpty := 0

	for row := startRow; row < startRow+size && row < len(grid); row++ {
		for col := startCol; col < startCol+size && col < len(grid[row]); col++ {
			total++
			if !grid[row][col].IsEmpty() {
				nonEmpty++
			}
		}
	}

	if total == 0 {
		return 0
	}
	return float64(nonEmpty) / float64(total)
}

// mergeOverlappingRegions merges overlapping table boundaries
func (ta *TableAnalyzer) mergeOverlappingRegions(regions []models.TableBoundary) []models.TableBoundary {
	if len(regions) <= 1 {
		return regions
	}

	merged := []models.TableBoundary{regions[0]}
	for i := 1; i < len(regions); i++ {
		current := regions[i]
		didMerge := false

		for j := range merged {
			if ta.overlaps(merged[j], current) {
				merged[j] = ta.merge(merged[j], current)
				didMerge = true
				break
			}
		}

		if !didMerge {
			merged = append(merged, current)
		}
	}

	return merged
}

// overlaps checks if two boundaries overlap
func (ta *TableAnalyzer) overlaps(a, b models.TableBoundary) bool {
	return !(a.EndRow < b.StartRow || b.EndRow < a.StartRow ||
		a.EndCol < b.StartCol || b.EndCol < a.StartCol)
}

// merge combines two boundaries into one
func (ta *TableAnalyzer) merge(a, b models.TableBoundary) models.TableBoundary {
	return models.TableBoundary{
		StartRow: minInt(a.StartRow, b.StartRow),
		EndRow:   maxInt(a.EndRow, b.EndRow),
		StartCol: minInt(a.StartCol, b.StartCol),
		EndCol:   maxInt(a.EndCol, b.EndCol),
	}
}

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}
