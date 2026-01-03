// Package models defines the core data structures for representing Excel data.
//
// This package contains all the types used throughout goxls for representing
// workbooks, sheets, tables, rows, and cells, as well as configuration and
// analysis structures.
//
// # Core Types
//
// The main types for representing Excel data:
//
//   - Workbook: Represents an entire Excel file
//   - Sheet: Represents a worksheet within a workbook
//   - Table: Represents a detected table within a sheet
//   - Row: Represents a data row with header-mapped values
//   - Cell: Represents a single cell with value, type, and metadata
//
// # Cell Types
//
// Cells can have different types:
//
//	const (
//	    CellTypeEmpty   // Empty cell
//	    CellTypeString  // Text value
//	    CellTypeNumber  // Numeric value
//	    CellTypeDate    // Date/time value
//	    CellTypeBool    // Boolean value
//	    CellTypeFormula // Formula (extracted as string)
//	)
//
// # Table Operations
//
// Tables support various operations:
//
//	// Filter rows
//	filtered := table.Filter(func(row Row) bool {
//	    if cell, ok := row.Get("Age"); ok {
//	        if val, ok := cell.AsFloat(); ok {
//	            return val > 18
//	        }
//	    }
//	    return false
//	})
//
//	// Column transformations
//	selected := table.Select("Name", "Email")
//	renamed := table.Rename(map[string]string{"old": "new"})
//	reordered := table.Reorder("Email", "Name")
//
//	// Deduplication
//	unique := table.Deduplicate("Email")
//	duplicates := table.FindDuplicates("Email")
//
//	// Column analysis
//	stats := table.AnalyzeColumns()
//
// # Table Comparison
//
// Compare two tables to find differences:
//
//	diff := models.DiffTables(oldTable, newTable, "ID")
//	if diff.HasChanges() {
//	    fmt.Printf("Added: %d, Removed: %d, Modified: %d\n",
//	        len(diff.AddedRows), len(diff.RemovedRows), len(diff.ModifiedRows))
//	}
//
// # Configuration
//
// Use DetectionConfig to customize table detection:
//
//	config := models.DefaultConfig()
//	config.MinColumns = 3
//	config.MinRows = 5
package models
