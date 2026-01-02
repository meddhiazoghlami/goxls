package models

import (
	"math"
	"time"
)

// CellType represents the type of data in a cell
type CellType int

const (
	CellTypeEmpty CellType = iota
	CellTypeString
	CellTypeNumber
	CellTypeDate
	CellTypeBool
	CellTypeFormula
)

// MergeRange represents a merged cell region
type MergeRange struct {
	StartRow int  // 0-indexed row of merge start
	StartCol int  // 0-indexed column of merge start
	EndRow   int  // 0-indexed row of merge end
	EndCol   int  // 0-indexed column of merge end
	IsOrigin bool // true if this cell is the top-left origin of the merge
}

// Cell represents a single cell in an Excel sheet
type Cell struct {
	Value        interface{}
	Type         CellType
	Row          int
	Col          int
	RawValue     string
	IsMerged     bool        // true if cell is part of a merged region
	MergeRange   *MergeRange // nil if not merged, points to merge info if merged
	Formula      string      // The formula string if cell contains a formula (e.g., "=SUM(A1:A10)")
	HasFormula   bool        // true if the cell contains a formula
	Comment      string      // Cell comment text (if any)
	HasComment   bool        // true if the cell has a comment
	Hyperlink    string      // Cell hyperlink URL (if any)
	HasHyperlink bool        // true if the cell has a hyperlink
}

// IsEmpty returns true if the cell is empty
func (c *Cell) IsEmpty() bool {
	return c.Type == CellTypeEmpty || c.RawValue == ""
}

// AsString returns the cell value as a string
func (c *Cell) AsString() string {
	if c.Value == nil {
		return ""
	}
	switch v := c.Value.(type) {
	case string:
		return v
	default:
		return c.RawValue
	}
}

// AsFloat returns the cell value as a float64
func (c *Cell) AsFloat() (float64, bool) {
	if v, ok := c.Value.(float64); ok {
		return v, true
	}
	return 0, false
}

// AsTime returns the cell value as a time.Time
func (c *Cell) AsTime() (time.Time, bool) {
	if v, ok := c.Value.(time.Time); ok {
		return v, true
	}
	return time.Time{}, false
}

// IsMergeOrigin returns true if this cell is the top-left origin of a merged region
func (c *Cell) IsMergeOrigin() bool {
	return c.MergeRange != nil && c.MergeRange.IsOrigin
}

// Row represents a data row with values mapped to headers
type Row struct {
	Index  int
	Values map[string]Cell
	Cells  []Cell
}

// Get returns the cell value for a given header
func (r *Row) Get(header string) (Cell, bool) {
	cell, ok := r.Values[header]
	return cell, ok
}

// Table represents a detected table within a sheet
type Table struct {
	Name       string
	Headers    []string
	Rows       []Row
	StartRow   int
	EndRow     int
	StartCol   int
	EndCol     int
	HeaderRow  int
}

// RowCount returns the number of data rows (excluding header)
func (t *Table) RowCount() int {
	return len(t.Rows)
}

// ColCount returns the number of columns
func (t *Table) ColCount() int {
	return len(t.Headers)
}

// RowPredicate is a function that evaluates a row and returns true if it matches
type RowPredicate func(row Row) bool

// Filter returns a new table containing only rows that match the predicate
func (t *Table) Filter(predicate RowPredicate) *Table {
	filtered := &Table{
		Name:      t.Name,
		Headers:   t.Headers,
		Rows:      make([]Row, 0),
		StartRow:  t.StartRow,
		EndRow:    t.EndRow,
		StartCol:  t.StartCol,
		EndCol:    t.EndCol,
		HeaderRow: t.HeaderRow,
	}

	for _, row := range t.Rows {
		if predicate(row) {
			filtered.Rows = append(filtered.Rows, row)
		}
	}

	return filtered
}

// FindDuplicates returns rows that have duplicate values in the specified key column
// The first occurrence is not included; only subsequent duplicates are returned
func (t *Table) FindDuplicates(keyColumn string) []Row {
	seen := make(map[string]bool)
	duplicates := make([]Row, 0)

	for _, row := range t.Rows {
		if cell, ok := row.Get(keyColumn); ok {
			key := cell.RawValue
			if seen[key] {
				duplicates = append(duplicates, row)
			} else {
				seen[key] = true
			}
		}
	}

	return duplicates
}

// Deduplicate returns a new table with duplicate rows removed based on the key column
// Keeps the first occurrence of each key value
func (t *Table) Deduplicate(keyColumn string) *Table {
	seen := make(map[string]bool)

	deduped := &Table{
		Name:      t.Name,
		Headers:   t.Headers,
		Rows:      make([]Row, 0),
		StartRow:  t.StartRow,
		EndRow:    t.EndRow,
		StartCol:  t.StartCol,
		EndCol:    t.EndCol,
		HeaderRow: t.HeaderRow,
	}

	for _, row := range t.Rows {
		if cell, ok := row.Get(keyColumn); ok {
			key := cell.RawValue
			if !seen[key] {
				seen[key] = true
				deduped.Rows = append(deduped.Rows, row)
			}
		}
	}

	return deduped
}

// DuplicateGroup represents a group of rows with the same key value
type DuplicateGroup struct {
	KeyValue string // The duplicate key value
	Rows     []Row  // All rows with this key value
	Count    int    // Number of occurrences
}

// FindDuplicateGroups returns groups of rows that share the same key value
// Only includes groups with more than one row (actual duplicates)
func (t *Table) FindDuplicateGroups(keyColumn string) []DuplicateGroup {
	groups := make(map[string][]Row)

	for _, row := range t.Rows {
		if cell, ok := row.Get(keyColumn); ok {
			key := cell.RawValue
			groups[key] = append(groups[key], row)
		}
	}

	// Filter to only groups with duplicates
	result := make([]DuplicateGroup, 0)
	for key, rows := range groups {
		if len(rows) > 1 {
			result = append(result, DuplicateGroup{
				KeyValue: key,
				Rows:     rows,
				Count:    len(rows),
			})
		}
	}

	return result
}

// Select returns a new table with only the specified columns
func (t *Table) Select(columns ...string) *Table {
	selected := &Table{
		Name:      t.Name,
		Headers:   make([]string, 0, len(columns)),
		Rows:      make([]Row, 0, len(t.Rows)),
		StartRow:  t.StartRow,
		EndRow:    t.EndRow,
		StartCol:  t.StartCol,
		EndCol:    t.EndCol,
		HeaderRow: t.HeaderRow,
	}

	// Build set of valid columns for quick lookup
	validColumns := make(map[string]bool)
	for _, h := range t.Headers {
		validColumns[h] = true
	}

	// Filter to only requested columns that exist
	for _, col := range columns {
		if validColumns[col] {
			selected.Headers = append(selected.Headers, col)
		}
	}

	// Copy rows with only selected columns
	for _, row := range t.Rows {
		newRow := Row{
			Index:  row.Index,
			Values: make(map[string]Cell),
			Cells:  make([]Cell, 0, len(selected.Headers)),
		}

		for _, col := range selected.Headers {
			if cell, ok := row.Values[col]; ok {
				newRow.Values[col] = cell
				newRow.Cells = append(newRow.Cells, cell)
			}
		}

		selected.Rows = append(selected.Rows, newRow)
	}

	return selected
}

// Rename returns a new table with columns renamed according to the mapping
// The map keys are old column names, values are new column names
func (t *Table) Rename(mapping map[string]string) *Table {
	renamed := &Table{
		Name:      t.Name,
		Headers:   make([]string, len(t.Headers)),
		Rows:      make([]Row, 0, len(t.Rows)),
		StartRow:  t.StartRow,
		EndRow:    t.EndRow,
		StartCol:  t.StartCol,
		EndCol:    t.EndCol,
		HeaderRow: t.HeaderRow,
	}

	// Rename headers
	for i, h := range t.Headers {
		if newName, ok := mapping[h]; ok {
			renamed.Headers[i] = newName
		} else {
			renamed.Headers[i] = h
		}
	}

	// Copy rows with renamed keys
	for _, row := range t.Rows {
		newRow := Row{
			Index:  row.Index,
			Values: make(map[string]Cell),
			Cells:  make([]Cell, len(row.Cells)),
		}
		copy(newRow.Cells, row.Cells)

		for oldName, cell := range row.Values {
			if newName, ok := mapping[oldName]; ok {
				newRow.Values[newName] = cell
			} else {
				newRow.Values[oldName] = cell
			}
		}

		renamed.Rows = append(renamed.Rows, newRow)
	}

	return renamed
}

// Reorder returns a new table with columns in the specified order
// Columns not in the list are excluded from the result
func (t *Table) Reorder(columns ...string) *Table {
	reordered := &Table{
		Name:      t.Name,
		Headers:   make([]string, 0, len(columns)),
		Rows:      make([]Row, 0, len(t.Rows)),
		StartRow:  t.StartRow,
		EndRow:    t.EndRow,
		StartCol:  t.StartCol,
		EndCol:    t.EndCol,
		HeaderRow: t.HeaderRow,
	}

	// Build set of valid columns
	validColumns := make(map[string]bool)
	for _, h := range t.Headers {
		validColumns[h] = true
	}

	// Keep only columns that exist, in the specified order
	for _, col := range columns {
		if validColumns[col] {
			reordered.Headers = append(reordered.Headers, col)
		}
	}

	// Copy rows with reordered columns
	for _, row := range t.Rows {
		newRow := Row{
			Index:  row.Index,
			Values: make(map[string]Cell),
			Cells:  make([]Cell, 0, len(reordered.Headers)),
		}

		for _, col := range reordered.Headers {
			if cell, ok := row.Values[col]; ok {
				newRow.Values[col] = cell
				newRow.Cells = append(newRow.Cells, cell)
			}
		}

		reordered.Rows = append(reordered.Rows, newRow)
	}

	return reordered
}

// ColumnStats holds statistical information about a column
type ColumnStats struct {
	Name          string   // Column header name
	Index         int      // Column index (0-based)
	InferredType  CellType // Most common non-empty cell type
	TotalCount    int      // Total number of cells
	EmptyCount    int      // Number of empty cells
	NullCount     int      // Number of null/empty values
	StringCount   int      // Number of string values
	NumberCount   int      // Number of numeric values
	DateCount     int      // Number of date values
	BoolCount     int      // Number of boolean values
	UniqueCount   int      // Number of unique values
	SampleValues  []string // Sample of values (up to 5)
	Min           float64  // Minimum numeric value (only valid if HasNumericStats is true)
	Max           float64  // Maximum numeric value (only valid if HasNumericStats is true)
	Sum           float64  // Sum of numeric values (only valid if HasNumericStats is true)
	Avg           float64  // Average of numeric values (only valid if HasNumericStats is true)
	HasNumericStats bool   // True if Min/Max/Sum/Avg are valid (column has numeric values)
}

// AnalyzeColumns returns statistical analysis for each column in the table
func (t *Table) AnalyzeColumns() []ColumnStats {
	if len(t.Headers) == 0 {
		return nil
	}

	stats := make([]ColumnStats, len(t.Headers))

	// Initialize stats for each column
	for i, header := range t.Headers {
		stats[i] = ColumnStats{
			Name:         header,
			Index:        i,
			TotalCount:   len(t.Rows),
			SampleValues: make([]string, 0, 5),
			Min:          math.MaxFloat64,
			Max:          -math.MaxFloat64,
		}
	}

	// Track unique values per column
	uniqueValues := make([]map[string]struct{}, len(t.Headers))
	for i := range uniqueValues {
		uniqueValues[i] = make(map[string]struct{})
	}

	// Analyze each row
	for _, row := range t.Rows {
		for i, header := range t.Headers {
			cell, exists := row.Values[header]
			if !exists || cell.IsEmpty() {
				stats[i].EmptyCount++
				stats[i].NullCount++
				continue
			}

			// Count by type
			switch cell.Type {
			case CellTypeString:
				stats[i].StringCount++
			case CellTypeNumber:
				stats[i].NumberCount++
				// Compute numeric statistics
				if val, ok := cell.AsFloat(); ok {
					stats[i].Sum += val
					if val < stats[i].Min {
						stats[i].Min = val
					}
					if val > stats[i].Max {
						stats[i].Max = val
					}
				}
			case CellTypeDate:
				stats[i].DateCount++
			case CellTypeBool:
				stats[i].BoolCount++
			case CellTypeEmpty:
				stats[i].EmptyCount++
				stats[i].NullCount++
			}

			// Track unique values
			rawVal := cell.RawValue
			if _, seen := uniqueValues[i][rawVal]; !seen {
				uniqueValues[i][rawVal] = struct{}{}

				// Collect sample values (up to 5)
				if len(stats[i].SampleValues) < 5 {
					stats[i].SampleValues = append(stats[i].SampleValues, rawVal)
				}
			}
		}
	}

	// Finalize stats
	for i := range stats {
		stats[i].UniqueCount = len(uniqueValues[i])
		stats[i].InferredType = inferColumnType(stats[i])

		// Compute average and set HasNumericStats flag
		if stats[i].NumberCount > 0 {
			stats[i].HasNumericStats = true
			stats[i].Avg = stats[i].Sum / float64(stats[i].NumberCount)
		} else {
			// Reset Min/Max to zero if no numeric values
			stats[i].Min = 0
			stats[i].Max = 0
		}
	}

	return stats
}

// inferColumnType determines the dominant type for a column
func inferColumnType(s ColumnStats) CellType {
	nonEmpty := s.TotalCount - s.EmptyCount
	if nonEmpty == 0 {
		return CellTypeEmpty
	}

	// Find the most common type
	maxCount := s.StringCount
	inferredType := CellTypeString

	if s.NumberCount > maxCount {
		maxCount = s.NumberCount
		inferredType = CellTypeNumber
	}
	if s.DateCount > maxCount {
		maxCount = s.DateCount
		inferredType = CellTypeDate
	}
	if s.BoolCount > maxCount {
		inferredType = CellTypeBool
	}

	return inferredType
}

// Sheet represents an Excel sheet containing one or more tables
type Sheet struct {
	Name   string
	Index  int
	Tables []Table
}

// Workbook represents an Excel file with multiple sheets
type Workbook struct {
	FilePath string
	Sheets   []Sheet
}

// TableBoundary represents the detected boundaries of a table
type TableBoundary struct {
	StartRow int
	EndRow   int
	StartCol int
	EndCol   int
}

// NamedRange represents an Excel named range
type NamedRange struct {
	Name     string // The name of the range (e.g., "SalesData")
	RefersTo string // The cell reference (e.g., "Sheet1!$A$1:$B$10")
	Scope    string // Either a sheet name or "Workbook" for global scope
}

// DetectionConfig holds configuration for table detection
type DetectionConfig struct {
	MinColumns         int     // Minimum columns to consider as a table
	MinRows            int     // Minimum rows to consider as a table
	MaxEmptyRows       int     // Max consecutive empty rows before table ends
	HeaderDensity      float64 // Minimum density of non-empty cells for header
	ColumnConsistency  float64 // Minimum consistency of column data types
	ExpandMergedCells  bool    // When true, copy merged cell value to all cells in range
	TrackMergeMetadata bool    // When true, populate IsMerged and MergeRange fields
}

// DefaultConfig returns the default detection configuration
func DefaultConfig() DetectionConfig {
	return DetectionConfig{
		MinColumns:         2,
		MinRows:            2,
		MaxEmptyRows:       2,
		HeaderDensity:      0.5,
		ColumnConsistency:  0.7,
		ExpandMergedCells:  true,
		TrackMergeMetadata: true,
	}
}

// CellDiff represents a change in a single cell
type CellDiff struct {
	Column   string // Column name where the change occurred
	OldValue string // Previous value (RawValue)
	NewValue string // New value (RawValue)
}

// RowDiff represents a modified row with its changes
type RowDiff struct {
	KeyValue string     // The key column value identifying this row
	OldRow   Row        // The original row
	NewRow   Row        // The modified row
	Changes  []CellDiff // List of cell changes
}

// DiffResult contains the differences between two tables
type DiffResult struct {
	AddedRows    []Row     // Rows present in new table but not in old
	RemovedRows  []Row     // Rows present in old table but not in new
	ModifiedRows []RowDiff // Rows present in both but with different values
	KeyColumn    string    // The column used as the key for comparison
}

// HasChanges returns true if there are any differences
func (d *DiffResult) HasChanges() bool {
	return len(d.AddedRows) > 0 || len(d.RemovedRows) > 0 || len(d.ModifiedRows) > 0
}

// TotalChanges returns the total number of changed rows
func (d *DiffResult) TotalChanges() int {
	return len(d.AddedRows) + len(d.RemovedRows) + len(d.ModifiedRows)
}

// DiffTables compares two tables and returns the differences
// The keyColumn is used to match rows between tables (like a primary key)
func DiffTables(oldTable, newTable *Table, keyColumn string) DiffResult {
	result := DiffResult{
		AddedRows:    make([]Row, 0),
		RemovedRows:  make([]Row, 0),
		ModifiedRows: make([]RowDiff, 0),
		KeyColumn:    keyColumn,
	}

	// Build a map of old rows by key
	oldRowMap := make(map[string]Row)
	for _, row := range oldTable.Rows {
		if cell, ok := row.Get(keyColumn); ok {
			key := cell.RawValue
			oldRowMap[key] = row
		}
	}

	// Build a map of new rows by key
	newRowMap := make(map[string]Row)
	for _, row := range newTable.Rows {
		if cell, ok := row.Get(keyColumn); ok {
			key := cell.RawValue
			newRowMap[key] = row
		}
	}

	// Find removed and modified rows
	for key, oldRow := range oldRowMap {
		if newRow, exists := newRowMap[key]; exists {
			// Row exists in both - check for modifications
			changes := compareRows(oldRow, newRow, oldTable.Headers, keyColumn)
			if len(changes) > 0 {
				result.ModifiedRows = append(result.ModifiedRows, RowDiff{
					KeyValue: key,
					OldRow:   oldRow,
					NewRow:   newRow,
					Changes:  changes,
				})
			}
		} else {
			// Row only in old table - it was removed
			result.RemovedRows = append(result.RemovedRows, oldRow)
		}
	}

	// Find added rows
	for key, newRow := range newRowMap {
		if _, exists := oldRowMap[key]; !exists {
			// Row only in new table - it was added
			result.AddedRows = append(result.AddedRows, newRow)
		}
	}

	return result
}

// compareRows compares two rows and returns the cell differences
func compareRows(oldRow, newRow Row, headers []string, keyColumn string) []CellDiff {
	changes := make([]CellDiff, 0)

	for _, header := range headers {
		// Skip the key column
		if header == keyColumn {
			continue
		}

		oldCell, oldExists := oldRow.Get(header)
		newCell, newExists := newRow.Get(header)

		oldValue := ""
		newValue := ""

		if oldExists {
			oldValue = oldCell.RawValue
		}
		if newExists {
			newValue = newCell.RawValue
		}

		if oldValue != newValue {
			changes = append(changes, CellDiff{
				Column:   header,
				OldValue: oldValue,
				NewValue: newValue,
			})
		}
	}

	return changes
}
