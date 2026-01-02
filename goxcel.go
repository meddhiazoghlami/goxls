// Package goxcel provides a lightweight, high-performance library for reading
// Excel files (.xlsx) with automatic table detection and intelligent data extraction.
//
// Basic usage:
//
//	workbook, err := goxcel.ReadFile("data.xlsx")
//	if err != nil {
//	    log.Fatal(err)
//	}
//	for _, sheet := range workbook.Sheets {
//	    for _, table := range sheet.Tables {
//	        fmt.Printf("Table: %s (%d rows)\n", table.Name, table.RowCount())
//	    }
//	}
//
// With options:
//
//	workbook, err := goxcel.ReadFile("data.xlsx",
//	    goxcel.WithMinColumns(3),
//	    goxcel.WithParallel(true),
//	)
//
// Export to JSON:
//
//	json, err := goxcel.ToJSON(table)
package goxcel

import (
	"context"
	"errors"
	"fmt"
	"os"

	"github.com/meddhiazoghlami/goxcel/pkg/export"
	"github.com/meddhiazoghlami/goxcel/pkg/models"
	"github.com/meddhiazoghlami/goxcel/pkg/reader"
)

// Re-export core types for convenience
type (
	// Workbook represents an entire Excel file
	Workbook = models.Workbook

	// Sheet represents a worksheet within a workbook
	Sheet = models.Sheet

	// Table represents a detected table within a sheet
	Table = models.Table

	// Row represents a data row with values mapped to headers
	Row = models.Row

	// Cell represents a single cell in an Excel sheet
	Cell = models.Cell

	// CellType represents the type of data in a cell
	CellType = models.CellType

	// MergeRange represents a merged cell region
	MergeRange = models.MergeRange

	// ColumnStats represents statistical analysis for a column
	ColumnStats = models.ColumnStats

	// DetectionConfig holds configuration for table detection
	DetectionConfig = models.DetectionConfig

	// DiffResult represents the result of comparing two tables
	DiffResult = models.DiffResult

	// RowDiff represents changes in a single row
	RowDiff = models.RowDiff

	// CellDiff represents a change in a single cell
	CellDiff = models.CellDiff

	// NamedRange represents an Excel named range
	NamedRange = models.NamedRange
)

// Re-export CellType constants
const (
	CellTypeEmpty   = models.CellTypeEmpty
	CellTypeString  = models.CellTypeString
	CellTypeNumber  = models.CellTypeNumber
	CellTypeDate    = models.CellTypeDate
	CellTypeBool    = models.CellTypeBool
	CellTypeFormula = models.CellTypeFormula
)

// Sentinel errors for common failure cases
var (
	// ErrFileNotFound is returned when the Excel file does not exist
	ErrFileNotFound = errors.New("goxcel: file not found")

	// ErrInvalidFormat is returned when the file is not a valid Excel file
	ErrInvalidFormat = errors.New("goxcel: invalid file format")

	// ErrSheetNotFound is returned when the requested sheet does not exist
	ErrSheetNotFound = errors.New("goxcel: sheet not found")

	// ErrNoTablesFound is returned when no tables are detected in the sheet
	ErrNoTablesFound = errors.New("goxcel: no tables detected")

	// ErrInvalidRange is returned when a cell range reference is invalid
	ErrInvalidRange = errors.New("goxcel: invalid cell range")

	// ErrEmptyWorkbook is returned when the workbook has no sheets
	ErrEmptyWorkbook = errors.New("goxcel: workbook is empty")

	// ErrContextCanceled is returned when the operation was canceled via context
	ErrContextCanceled = errors.New("goxcel: operation canceled")
)

// Option is a functional option for configuring the reader
type Option func(*options)

// options holds all configuration for reading
type options struct {
	config   models.DetectionConfig
	parallel bool
}

// defaultOptions returns the default options
func defaultOptions() *options {
	return &options{
		config:   models.DefaultConfig(),
		parallel: false,
	}
}

// WithMinColumns sets the minimum number of columns for table detection
func WithMinColumns(n int) Option {
	return func(o *options) {
		o.config.MinColumns = n
	}
}

// WithMinRows sets the minimum number of rows for table detection
func WithMinRows(n int) Option {
	return func(o *options) {
		o.config.MinRows = n
	}
}

// WithMaxEmptyRows sets the maximum consecutive empty rows before table ends
func WithMaxEmptyRows(n int) Option {
	return func(o *options) {
		o.config.MaxEmptyRows = n
	}
}

// WithHeaderDensity sets the minimum density of non-empty cells for header detection
func WithHeaderDensity(d float64) Option {
	return func(o *options) {
		o.config.HeaderDensity = d
	}
}

// WithColumnConsistency sets the minimum consistency of column data types
func WithColumnConsistency(c float64) Option {
	return func(o *options) {
		o.config.ColumnConsistency = c
	}
}

// WithExpandMergedCells enables/disables copying merged cell values to all cells in range
func WithExpandMergedCells(expand bool) Option {
	return func(o *options) {
		o.config.ExpandMergedCells = expand
	}
}

// WithTrackMergeMetadata enables/disables populating IsMerged and MergeRange fields
func WithTrackMergeMetadata(track bool) Option {
	return func(o *options) {
		o.config.TrackMergeMetadata = track
	}
}

// WithParallel enables/disables parallel sheet processing
func WithParallel(parallel bool) Option {
	return func(o *options) {
		o.parallel = parallel
	}
}

// WithConfig sets the full detection configuration
func WithConfig(config DetectionConfig) Option {
	return func(o *options) {
		o.config = config
	}
}

// ReadFile reads an Excel file and extracts all tables from all sheets.
// It accepts optional configuration via functional options.
//
// Example:
//
//	workbook, err := goxcel.ReadFile("data.xlsx")
//	workbook, err := goxcel.ReadFile("data.xlsx", goxcel.WithParallel(true))
//	workbook, err := goxcel.ReadFile("data.xlsx", goxcel.WithMinColumns(3), goxcel.WithMinRows(5))
func ReadFile(filePath string, opts ...Option) (*Workbook, error) {
	// Check if file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, fmt.Errorf("%w: %s", ErrFileNotFound, filePath)
	}

	// Apply options
	o := defaultOptions()
	for _, opt := range opts {
		opt(o)
	}

	// Create reader with config
	wr := reader.NewWorkbookReaderWithConfig(o.config)

	// Read file
	var workbook *Workbook
	var err error

	if o.parallel {
		workbook, err = wr.ReadFileParallel(filePath)
	} else {
		workbook, err = wr.ReadFile(filePath)
	}

	if err != nil {
		return nil, wrapError(err)
	}

	return workbook, nil
}

// ReadFileWithContext reads an Excel file with context support for cancellation.
// The context can be used to cancel long-running operations.
//
// Example:
//
//	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
//	defer cancel()
//	workbook, err := goxcel.ReadFileWithContext(ctx, "large.xlsx")
func ReadFileWithContext(ctx context.Context, filePath string, opts ...Option) (*Workbook, error) {
	// Check context before starting
	select {
	case <-ctx.Done():
		return nil, fmt.Errorf("%w: %v", ErrContextCanceled, ctx.Err())
	default:
	}

	// Create a channel for the result
	type result struct {
		workbook *Workbook
		err      error
	}
	done := make(chan result, 1)

	// Run ReadFile in a goroutine
	go func() {
		wb, err := ReadFile(filePath, opts...)
		done <- result{wb, err}
	}()

	// Wait for either completion or context cancellation
	select {
	case <-ctx.Done():
		return nil, fmt.Errorf("%w: %v", ErrContextCanceled, ctx.Err())
	case r := <-done:
		return r.workbook, r.err
	}
}

// ReadSheet reads a specific sheet from an Excel file.
//
// Example:
//
//	sheet, err := goxcel.ReadSheet("data.xlsx", "Sales")
func ReadSheet(filePath, sheetName string, opts ...Option) (*Sheet, error) {
	// Check if file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, fmt.Errorf("%w: %s", ErrFileNotFound, filePath)
	}

	// Apply options
	o := defaultOptions()
	for _, opt := range opts {
		opt(o)
	}

	// Create reader with config
	wr := reader.NewWorkbookReaderWithConfig(o.config)

	// Read sheet
	sheet, err := wr.ReadSheet(filePath, sheetName)
	if err != nil {
		// Check if it's a "sheet not found" error
		if sheet == nil {
			return nil, fmt.Errorf("%w: %s", ErrSheetNotFound, sheetName)
		}
		return nil, wrapError(err)
	}

	return sheet, nil
}

// DefaultConfig returns the default detection configuration
func DefaultConfig() DetectionConfig {
	return models.DefaultConfig()
}

// DiffTables compares two tables and returns the differences.
// The keyColumn is used to match rows between tables.
//
// Example:
//
//	diff := goxcel.DiffTables(oldTable, newTable, "ID")
//	if diff.HasChanges() {
//	    fmt.Printf("Added: %d, Removed: %d, Modified: %d\n",
//	        len(diff.AddedRows), len(diff.RemovedRows), len(diff.ModifiedRows))
//	}
func DiffTables(oldTable, newTable *Table, keyColumn string) DiffResult {
	return models.DiffTables(oldTable, newTable, keyColumn)
}

// --- Export Functions ---

// ToJSON exports a table to JSON format.
//
// Example:
//
//	json, err := goxcel.ToJSON(table)
func ToJSON(table *Table) (string, error) {
	return export.ToJSON(table)
}

// ToJSONPretty exports a table to pretty-printed JSON format.
//
// Example:
//
//	json, err := goxcel.ToJSONPretty(table)
func ToJSONPretty(table *Table) (string, error) {
	return export.ToJSONPretty(table)
}

// ToCSV exports a table to CSV format.
//
// Example:
//
//	csv, err := goxcel.ToCSV(table)
func ToCSV(table *Table) (string, error) {
	return export.ToCSV(table)
}

// ToTSV exports a table to TSV (tab-separated values) format.
//
// Example:
//
//	tsv, err := goxcel.ToTSV(table)
func ToTSV(table *Table) (string, error) {
	return export.ToTSV(table)
}

// ToCSVWithDelimiter exports a table to CSV format with a custom delimiter.
//
// Example:
//
//	csv, err := goxcel.ToCSVWithDelimiter(table, ';')
func ToCSVWithDelimiter(table *Table, delimiter rune) (string, error) {
	return export.ToCSVWithDelimiter(table, delimiter)
}

// ToSQL exports a table to SQL INSERT statements.
//
// Example:
//
//	sql, err := goxcel.ToSQL(table, "users")
func ToSQL(table *Table, tableName string) (string, error) {
	return export.ToSQL(table, tableName)
}

// ToSQLWithCreate exports a table to SQL with CREATE TABLE statement.
//
// Example:
//
//	sql, err := goxcel.ToSQLWithCreate(table, "users")
func ToSQLWithCreate(table *Table, tableName string) (string, error) {
	return export.ToSQLWithCreate(table, tableName)
}

// wrapError wraps internal errors with sentinel errors where appropriate
func wrapError(err error) error {
	if err == nil {
		return nil
	}

	errStr := err.Error()

	// Check for common error patterns
	if contains(errStr, "no such file") || contains(errStr, "file not found") {
		return fmt.Errorf("%w: %v", ErrFileNotFound, err)
	}
	if contains(errStr, "invalid") || contains(errStr, "zip") {
		return fmt.Errorf("%w: %v", ErrInvalidFormat, err)
	}
	if contains(errStr, "sheet") && contains(errStr, "not found") {
		return fmt.Errorf("%w: %v", ErrSheetNotFound, err)
	}

	return err
}

// contains checks if s contains substr (case-insensitive would be better but keeping simple)
func contains(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || len(s) > 0 && containsHelper(s, substr))
}

func containsHelper(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}
