package reader

import (
	"fmt"
	"strings"

	"excel-lite/pkg/models"
)

// HeaderDetector identifies header rows in tables
type HeaderDetector struct {
	config models.DetectionConfig
}

// NewHeaderDetector creates a new header detector
func NewHeaderDetector(config models.DetectionConfig) *HeaderDetector {
	return &HeaderDetector{config: config}
}

// NewDefaultHeaderDetector creates a header detector with default config
func NewDefaultHeaderDetector() *HeaderDetector {
	return NewHeaderDetector(models.DefaultConfig())
}

// DetectHeaderRow finds the most likely header row within a boundary
func (hd *HeaderDetector) DetectHeaderRow(grid [][]models.Cell, boundary models.TableBoundary) int {
	bestRow := boundary.StartRow
	bestScore := 0.0

	// Check each row in the boundary for header characteristics
	for row := boundary.StartRow; row <= minInt(boundary.StartRow+5, boundary.EndRow); row++ {
		score := hd.scoreAsHeader(grid, row, boundary)
		if score > bestScore {
			bestScore = score
			bestRow = row
		}
	}

	return bestRow
}

// scoreAsHeader scores how likely a row is to be a header
func (hd *HeaderDetector) scoreAsHeader(grid [][]models.Cell, row int, boundary models.TableBoundary) float64 {
	if row >= len(grid) {
		return 0
	}

	score := 0.0
	cellCount := 0
	stringCount := 0
	nonEmptyCount := 0

	for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[row]); col++ {
		cell := grid[row][col]
		cellCount++

		if !cell.IsEmpty() {
			nonEmptyCount++
			// Headers are typically strings
			if cell.Type == models.CellTypeString {
				stringCount++
			}
		}
	}

	if cellCount == 0 {
		return 0
	}

	// Score based on density of non-empty cells
	density := float64(nonEmptyCount) / float64(cellCount)
	score += density * 40

	// Score based on proportion of string values
	if nonEmptyCount > 0 {
		stringRatio := float64(stringCount) / float64(nonEmptyCount)
		score += stringRatio * 30
	}

	// Bonus if row has more strings than the next row (data rows often have mixed types)
	if row+1 <= boundary.EndRow && row+1 < len(grid) {
		nextRowStrings := hd.countStrings(grid, row+1, boundary)
		if stringCount > nextRowStrings {
			score += 20
		}
	}

	// Bonus if headers look like typical header patterns
	score += hd.patternBonus(grid, row, boundary)

	// Bonus for merged cells (common in headers for grouping columns)
	mergedCount := 0
	for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[row]); col++ {
		cell := grid[row][col]
		if cell.IsMerged && cell.MergeRange != nil && cell.MergeRange.IsOrigin {
			mergedCount++
		}
	}
	if mergedCount > 0 {
		score += float64(mergedCount) * 5
	}

	return score
}

// countStrings counts string-type cells in a row
func (hd *HeaderDetector) countStrings(grid [][]models.Cell, row int, boundary models.TableBoundary) int {
	if row >= len(grid) {
		return 0
	}

	count := 0
	for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[row]); col++ {
		if grid[row][col].Type == models.CellTypeString {
			count++
		}
	}
	return count
}

// patternBonus gives bonus points for common header patterns
func (hd *HeaderDetector) patternBonus(grid [][]models.Cell, row int, boundary models.TableBoundary) float64 {
	if row >= len(grid) {
		return 0
	}

	bonus := 0.0
	commonHeaders := []string{
		"id", "name", "date", "time", "type", "status", "email", "phone",
		"address", "city", "state", "country", "zip", "code", "description",
		"price", "amount", "quantity", "total", "number", "count", "value",
		"first", "last", "created", "updated", "modified", "title", "category",
	}

	for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[row]); col++ {
		cell := grid[row][col]
		if cell.IsEmpty() {
			continue
		}

		value := strings.ToLower(strings.TrimSpace(cell.AsString()))
		for _, header := range commonHeaders {
			if strings.Contains(value, header) {
				bonus += 2
				break
			}
		}
	}

	return bonus
}

// ExtractHeaders extracts header names from a row
func (hd *HeaderDetector) ExtractHeaders(grid [][]models.Cell, headerRow int, boundary models.TableBoundary) []string {
	if headerRow >= len(grid) {
		return nil
	}

	headers := make([]string, 0, boundary.EndCol-boundary.StartCol+1)
	usedNames := make(map[string]int)

	for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[headerRow]); col++ {
		cell := grid[headerRow][col]
		header := hd.normalizeHeader(cell.AsString(), col, usedNames)
		headers = append(headers, header)
	}

	return headers
}

// normalizeHeader cleans and ensures uniqueness of header names
func (hd *HeaderDetector) normalizeHeader(value string, colIndex int, usedNames map[string]int) string {
	// Trim and clean
	header := strings.TrimSpace(value)

	// If empty, generate a default name
	if header == "" {
		header = fmt.Sprintf("Column_%d", colIndex+1)
	}

	// Handle duplicates by appending a number
	originalHeader := header
	count, exists := usedNames[strings.ToLower(header)]
	if exists {
		header = fmt.Sprintf("%s_%d", originalHeader, count+1)
	}
	usedNames[strings.ToLower(originalHeader)] = count + 1

	return header
}

// ValidateHeaders checks if detected headers are reasonable
func (hd *HeaderDetector) ValidateHeaders(headers []string) bool {
	if len(headers) < hd.config.MinColumns {
		return false
	}

	// Count non-empty, non-generic headers
	meaningfulCount := 0
	for _, h := range headers {
		if !strings.HasPrefix(h, "Column_") && h != "" {
			meaningfulCount++
		}
	}

	// At least half should be meaningful
	return float64(meaningfulCount)/float64(len(headers)) >= 0.5
}

// DetectHeaderRows detects multi-row headers (common with merged cells)
// Returns the start and end row indices for the header region
func (hd *HeaderDetector) DetectHeaderRows(grid [][]models.Cell, boundary models.TableBoundary) (headerStart int, headerEnd int) {
	headerStart = boundary.StartRow
	headerEnd = boundary.StartRow

	if headerStart >= len(grid) {
		return headerStart, headerEnd
	}

	// Check if first row has merged cells spanning multiple rows
	for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[headerStart]); col++ {
		cell := grid[headerStart][col]
		if cell.MergeRange != nil && cell.MergeRange.EndRow > headerEnd {
			// Merged cell spans multiple rows - extend header range
			headerEnd = cell.MergeRange.EndRow
		}
	}

	// Limit header rows to a reasonable maximum
	maxHeaderRows := 3
	if headerEnd-headerStart >= maxHeaderRows {
		headerEnd = headerStart + maxHeaderRows - 1
	}

	// Ensure we don't exceed the table boundary
	if headerEnd > boundary.EndRow {
		headerEnd = boundary.EndRow
	}

	return headerStart, headerEnd
}

// ExtractHierarchicalHeaders extracts multi-level header structure for merged headers
// Returns a 2D slice where each inner slice represents one header level
func (hd *HeaderDetector) ExtractHierarchicalHeaders(grid [][]models.Cell, headerStart, headerEnd int, boundary models.TableBoundary) [][]string {
	levels := headerEnd - headerStart + 1
	if levels <= 0 {
		return nil
	}

	result := make([][]string, levels)

	for level := 0; level < levels; level++ {
		row := headerStart + level
		if row >= len(grid) {
			continue
		}

		result[level] = make([]string, 0, boundary.EndCol-boundary.StartCol+1)
		usedNames := make(map[string]int)

		for col := boundary.StartCol; col <= boundary.EndCol && col < len(grid[row]); col++ {
			cell := grid[row][col]
			header := hd.normalizeHeader(cell.AsString(), col, usedNames)
			result[level] = append(result[level], header)
		}
	}

	return result
}

// FlattenHierarchicalHeaders combines multi-level headers into single strings
// using the given separator (e.g., " > " to get "Group > Subgroup")
func (hd *HeaderDetector) FlattenHierarchicalHeaders(hierarchical [][]string, separator string) []string {
	if len(hierarchical) == 0 {
		return nil
	}

	// Use the last level as the base (most specific headers)
	lastLevel := hierarchical[len(hierarchical)-1]
	result := make([]string, len(lastLevel))

	for col := range lastLevel {
		parts := make([]string, 0, len(hierarchical))

		// Collect non-empty headers from each level for this column
		for level := 0; level < len(hierarchical); level++ {
			if col < len(hierarchical[level]) {
				header := hierarchical[level][col]
				// Skip generic column names and empty strings
				if header != "" && !strings.HasPrefix(header, "Column_") {
					// Avoid duplicating the same header from merged cells
					if len(parts) == 0 || parts[len(parts)-1] != header {
						parts = append(parts, header)
					}
				}
			}
		}

		if len(parts) > 0 {
			result[col] = strings.Join(parts, separator)
		} else {
			result[col] = fmt.Sprintf("Column_%d", col+1)
		}
	}

	return result
}
