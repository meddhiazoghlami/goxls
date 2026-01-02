# Excel-Lite

A lightweight, high-performance Go library for reading Excel files (.xlsx) with automatic table detection and intelligent data extraction.

![Go Version](https://img.shields.io/badge/Go-1.25+-00ADD8?style=flat&logo=go)
![Coverage](https://img.shields.io/badge/coverage-95%25-brightgreen)


## Overview

Excel-Lite automatically detects and extracts tabular data from Excel spreadsheets without requiring predefined schemas or manual range specifications. It intelligently identifies table boundaries, headers, and data types, making it ideal for processing Excel files with unknown or dynamic structures.

## Features

- **Automatic Table Detection** - Identifies table boundaries using density analysis and pattern recognition
- **Smart Header Detection** - Automatically identifies header rows using scoring algorithms
- **Type Inference** - Detects cell types: strings, numbers, dates, booleans, and formulas
- **Formula Extraction** - Extracts formula strings from cells (e.g., `=SUM(A1:A10)`)
- **Cell Comments Support** - Reads cell comments/notes from Excel files
- **Hyperlink Support** - Extracts hyperlinks from cells (URLs, mailto, internal references)
- **Merged Cell Support** - Detects merged regions, expands values, and tracks merge metadata
- **Multi-Table Support** - Extracts multiple tables from a single sheet
- **Concurrent Processing** - Process multiple sheets in parallel for better performance
- **Named Range Support** - Read Excel named ranges and extract them as tables
- **Data Validation** - Validate table data against customizable rules
- **Excel Date Handling** - Converts Excel serial dates to Go `time.Time` (handles the 1900 leap year bug)
- **Multiple Export Formats** - Export to JSON, CSV, or SQL with configurable options
- **SQL Dialect Support** - Generate SQL for MySQL, PostgreSQL, SQLite, or generic SQL
- **Configurable Detection** - Fine-tune table detection parameters for edge cases
- **High Test Coverage** - 95%+ test coverage across all packages

## Installation

```bash
go get github.com/yourusername/excel-lite
```

## Quick Start

### Basic Usage

```go
package main

import (
    "fmt"
    "excel-lite/pkg/reader"
)

func main() {
    // Create a workbook reader
    wr := reader.NewWorkbookReader()

    // Read an Excel file
    workbook, err := wr.ReadFile("data.xlsx")
    if err != nil {
        panic(err)
    }

    // Iterate through sheets and tables
    for _, sheet := range workbook.Sheets {
        fmt.Printf("Sheet: %s\n", sheet.Name)

        for _, table := range sheet.Tables {
            fmt.Printf("  Table: %s (%d rows)\n", table.Name, table.RowCount())
            fmt.Printf("  Headers: %v\n", table.Headers)

            // Access row data
            for _, row := range table.Rows {
                if cell, ok := row.Get("Name"); ok {
                    fmt.Printf("    Name: %s\n", cell.AsString())
                }
            }
        }
    }
}
```

### Export to JSON

```go
import "excel-lite/pkg/export"

// Simple export
jsonStr, err := export.ToJSON(table)

// With options
opts := export.DefaultJSONOptions()
opts.Pretty = true
opts.SelectedColumns = []string{"Name", "Email", "Age"}

exporter := export.NewJSONExporter(opts)
result, err := exporter.ExportString(table)
```

### Export to CSV

```go
import "excel-lite/pkg/export"

// Simple export
csvStr, err := export.ToCSV(table)

// With options
opts := export.DefaultCSVOptions()
opts.Delimiter = ';'
opts.QuoteAll = true

exporter := export.NewCSVExporter(opts)
result, err := exporter.ExportString(table)
```

### Export to SQL

```go
import "excel-lite/pkg/export"

// Simple export
sqlStr, err := export.ToSQL(table, "users")

// With options (PostgreSQL with CREATE TABLE)
opts := export.DefaultSQLOptions()
opts.TableName = "employees"
opts.Dialect = export.DialectPostgreSQL
opts.CreateTable = true
opts.BatchSize = 100

exporter := export.NewSQLExporter(opts)
result, err := exporter.ExportString(table)
```

### Excel Date Conversion

```go
import "excel-lite/pkg/dateutil"

// Convert Excel serial date to Go time.Time
// Serial 45658 = January 1, 2025
t := dateutil.ExcelDateToTime(45658)

// Convert Go time back to Excel serial
serial := dateutil.TimeToExcelDate(time.Now())

// Format Excel date directly
formatted := dateutil.FormatExcelDate(45658, "2006-01-02")
// Output: "2025-01-01"

// Convert with specific timezone
loc, _ := time.LoadLocation("America/New_York")
t = dateutil.ExcelDateToTimeWithLocation(45658, loc)
```

## API Reference

### Package `reader`

The main entry point for reading Excel files.

```go
// Create a reader with default configuration
wr := reader.NewWorkbookReader()

// Create a reader with custom configuration
config := models.DetectionConfig{
    MinColumns:         2,     // Minimum columns for table detection
    MinRows:            2,     // Minimum rows for table detection
    MaxEmptyRows:       2,     // Max empty rows before table boundary
    HeaderDensity:      0.5,   // Min density for header row
    ColumnConsistency:  0.7,   // Min type consistency for columns
    ExpandMergedCells:  true,  // Copy merged cell value to all cells in range
    TrackMergeMetadata: true,  // Populate IsMerged and MergeRange fields
}
wr := reader.NewWorkbookReaderWithConfig(config)

// Read entire workbook
workbook, err := wr.ReadFile("file.xlsx")

// Read with concurrent sheet processing (faster for multi-sheet workbooks)
workbook, err := wr.ReadFileParallel("file.xlsx")

// Read specific sheet
sheet, err := wr.ReadSheet("file.xlsx", "Sheet1")

// Utility functions
table := reader.GetTableByName(workbook, "Sheet1_Table1")
tables := reader.GetAllTables(workbook)
```

### Package `models`

Core data structures for representing Excel data.

```go
// Workbook - represents an entire Excel file
type Workbook struct {
    FilePath string
    Sheets   []Sheet
}

// Sheet - represents a worksheet
type Sheet struct {
    Name   string
    Index  int
    Tables []Table
}

// Table - represents a detected table
type Table struct {
    Name      string
    Headers   []string
    Rows      []Row
    StartRow  int
    EndRow    int
    StartCol  int
    EndCol    int
    HeaderRow int
}

// Row - represents a data row
type Row struct {
    Index  int
    Values map[string]Cell
    Cells  []Cell
}

// Cell - represents a single cell
type Cell struct {
    Value        interface{}
    Type         CellType     // Empty, String, Number, Date, Bool, Formula
    Row          int
    Col          int
    RawValue     string
    IsMerged     bool         // true if part of a merged region
    MergeRange   *MergeRange  // merge info (nil if not merged)
    Formula      string       // formula string if cell contains a formula
    HasFormula   bool         // true if cell contains a formula
    Comment      string       // cell comment text (if any)
    HasComment   bool         // true if cell has a comment
    Hyperlink    string       // hyperlink URL (if any)
    HasHyperlink bool         // true if cell has a hyperlink
}

// MergeRange - represents a merged cell region
type MergeRange struct {
    StartRow int
    StartCol int
    EndRow   int
    EndCol   int
    IsOrigin bool  // true if this cell is the top-left of the merge
}

// ColumnStats - statistical analysis for a column
type ColumnStats struct {
    Name            string
    Index           int
    InferredType    CellType
    TotalCount      int
    EmptyCount      int
    UniqueCount     int
    StringCount     int
    NumberCount     int
    DateCount       int
    BoolCount       int
    SampleValues    []string  // Up to 5 unique values
    Min             float64   // Minimum (only if HasNumericStats)
    Max             float64   // Maximum (only if HasNumericStats)
    Sum             float64   // Sum (only if HasNumericStats)
    Avg             float64   // Average (only if HasNumericStats)
    HasNumericStats bool      // True if numeric stats are valid
}

// Get column statistics
stats := table.AnalyzeColumns()
for _, col := range stats {
    fmt.Printf("Column: %s, Type: %v, Unique: %d\n", col.Name, col.InferredType, col.UniqueCount)
    if col.HasNumericStats {
        fmt.Printf("  Min: %.2f, Max: %.2f, Avg: %.2f\n", col.Min, col.Max, col.Avg)
    }
}

// Filter rows using a predicate function
filtered := table.Filter(func(row Row) bool {
    if cell, ok := row.Get("Age"); ok {
        if val, ok := cell.AsFloat(); ok {
            return val > 18
        }
    }
    return false
})

// Chain filters for complex queries
active := table.Filter(isActive).Filter(hasValidEmail)

// Compare two tables and find differences
result := DiffTables(oldTable, newTable, "ID") // Use "ID" as key column

if result.HasChanges() {
    fmt.Printf("Added: %d, Removed: %d, Modified: %d\n",
        len(result.AddedRows), len(result.RemovedRows), len(result.ModifiedRows))
}

// Inspect what changed in modified rows
for _, mod := range result.ModifiedRows {
    for _, change := range mod.Changes {
        fmt.Printf("%s: %s -> %s\n", change.Column, change.OldValue, change.NewValue)
    }
}

// Find and remove duplicate rows
duplicates := table.FindDuplicates("Email")      // Get duplicate rows
unique := table.Deduplicate("Email")             // Remove duplicates
groups := table.FindDuplicateGroups("Email")     // Get groups with counts

// Column transformations
selected := table.Select("Name", "Email")        // Keep only these columns
renamed := table.Rename(map[string]string{"old": "new"})  // Rename columns
reordered := table.Reorder("Email", "Name")      // Change column order

// Chain transformations
result := table.Select("name", "email").Rename(map[string]string{"name": "Name"})
```

### Package `export`

Export tables to various formats.

| Function | Description |
|----------|-------------|
| `ToJSON(table)` | Export to JSON with defaults |
| `ToCSV(table)` | Export to CSV with defaults |
| `ToSQL(table, tableName)` | Export to SQL INSERT statements |
| `NewExporter(format, opts)` | Create exporter with custom options |

**Supported Formats:**
- `FormatJSON` - JSON array of objects
- `FormatCSV` - Comma-separated values
- `FormatSQL` - SQL INSERT statements

**SQL Dialects:**
- `DialectGeneric` - Standard SQL
- `DialectMySQL` - MySQL-specific syntax
- `DialectPostgreSQL` - PostgreSQL-specific syntax
- `DialectSQLite` - SQLite-specific syntax

### Package `dateutil`

Excel date conversion utilities.

| Function | Description |
|----------|-------------|
| `ExcelDateToTime(serial)` | Convert Excel serial to `time.Time` |
| `TimeToExcelDate(t)` | Convert `time.Time` to Excel serial |
| `IsExcelDateSerial(value)` | Check if value is likely an Excel date |
| `ExcelDateToTimeWithLocation(serial, loc)` | Convert with timezone |
| `FormatExcelDate(serial, layout)` | Convert and format in one call |

### Package `validation`

Validate table data against defined rules.

```go
import "excel-lite/pkg/validation"

// Using the fluent RuleBuilder API
rules := []validation.ValidationRule{
    validation.ForColumn("Email").Required().MatchesPattern(`^[\w.-]+@[\w.-]+\.\w+$`).Build(),
    validation.ForColumn("Age").Range(18, 120).Build(),
    validation.ForColumn("Status").OneOf("active", "inactive", "pending").Build(),
}

// Or using struct directly
rules := []validation.ValidationRule{
    {Column: "Name", Required: true},
    {Column: "Price", MinVal: 0, MinValSet: true},
}

// Validate a table
result := validation.ValidateTable(table, rules)
if !result.Valid {
    for _, err := range result.Errors {
        fmt.Printf("Row %d, %s: %s\n", err.Row, err.Column, err.Message)
    }
}

// Group errors for analysis
byColumn := result.ErrorsByColumn()
byRow := result.ErrorsByRow()
```

**Validation Rule Options:**
- `Required` - Field cannot be empty
- `Pattern` - Value must match regex pattern
- `MinVal/MaxVal` - Numeric range validation
- `AllowedValues` - Value must be in allowed list
- `CustomFunc` - Custom validation function

### Named Ranges

Read Excel named ranges directly as tables.

```go
import "excel-lite/pkg/reader"

// Create a named range reader
nr := reader.NewNamedRangeReader()

// List all named ranges in a file
ranges, _ := nr.GetNamedRanges("data.xlsx")
for _, r := range ranges {
    fmt.Printf("Name: %s, RefersTo: %s, Scope: %s\n", r.Name, r.RefersTo, r.Scope)
}

// Read a named range as a table
table, _ := nr.ReadRange("data.xlsx", "SalesData")
fmt.Printf("Table: %s with %d rows\n", table.Name, len(table.Rows))

// Helper functions
r := reader.GetNamedRangeByName(ranges, "MyRange")
global := reader.GetGlobalNamedRanges(ranges)        // Workbook-scoped ranges
byScope := reader.GetNamedRangesByScope(ranges, "Sheet1")  // Sheet-scoped ranges

// Parse range reference info
info, _ := reader.ParseNamedRangeInfo("Sheet1!$A$1:$B$10")
fmt.Printf("Sheet: %s, Start: %s, End: %s\n", info.SheetName, info.StartCell, info.EndCell)
```

## CLI Usage

```bash
# Build the CLI
go build -o excel-lite ./cmd/main.go

# Read and analyze an Excel file
./excel-lite data.xlsx

# Export to JSON (pretty printed)
./excel-lite -f json --pretty data.xlsx

# Export to CSV
./excel-lite -f csv -o output.csv data.xlsx

# Export to SQL
./excel-lite -f sql --sql-table=users data.xlsx

# Filter by sheet name
./excel-lite -s Sales data.xlsx

# Filter by table name
./excel-lite -t Sales_Table1 data.xlsx

# Select specific columns
./excel-lite -f csv -c "Name,Email,Age" data.xlsx

# Quick summary with column type analysis
./excel-lite --summary data.xlsx
```

### CLI Options

| Option | Short | Description |
|--------|-------|-------------|
| `--format` | `-f` | Output format: json, csv, sql, text (default: text) |
| `--output` | `-o` | Output file path (default: stdout) |
| `--sheet` | `-s` | Filter by sheet name |
| `--table` | `-t` | Filter by table name |
| `--columns` | `-c` | Comma-separated columns to include |
| `--sql-table` | | Table name for SQL output (default: data) |
| `--summary` | | Show summary with column type analysis |
| `--pretty` | | Pretty print JSON output |
| `--no-headers` | | Exclude headers from CSV output |

### Example Output

```
=== Workbook Summary ===
File: data.xlsx
Sheets: 2
Total Tables Detected: 3

=== Sheet: Sales ===

  Table: Sales_Table1
  Location: Row 1-150, Col 1-5
  Header Row: 1
  Columns: 5
  Rows: 149
  Headers: [Date Product Quantity Price Total]
  Sample Data (first 3 rows):
    Row 2: Date="2025-01-01", Product="Widget", Quantity="10", Price="25.00", Total="250.00"
    ...
```

## Known Limitations

- **No .xls Support** - Only .xlsx format (Office 2007+) is supported
- **Large Files** - Entire file is loaded into memory; not suitable for very large files (100k+ rows)
- **No Write Support** - Read-only; cannot create or modify Excel files
- **No Formula Evaluation** - Formulas are extracted as strings but not evaluated

## Test Coverage

| Package | Coverage | Tests |
|---------|----------|-------|
| `pkg/dateutil` | 100% | 10 test functions |
| `pkg/models` | 96.8% | Full coverage |
| `pkg/reader` | 96.9% | 350+ tests |
| `pkg/export` | 95.4% | 60+ tests |
| `pkg/validation` | 95.1% | 25+ tests |
| **Total** | **95%+** | 450+ tests |

Run tests:
```bash
go test ./... -v
go test ./... -cover
```

## Roadmap

### Completed
- [x] Merged cell support (value expansion, metadata tracking, multi-row headers)
- [x] Column type inference and statistics (`table.AnalyzeColumns()`) with Min/Max/Sum/Avg
- [x] Row filtering (`table.Filter()`) with chainable predicates
- [x] Table diff (`DiffTables()`) to compare tables and find changes
- [x] Row deduplication (`FindDuplicates()`, `Deduplicate()`)
- [x] Column transformations (`Select()`, `Rename()`, `Reorder()`)
- [x] Performance benchmarks
- [x] CLI enhancements (output formats, filtering, column selection)
- [x] Formula extraction (read formula strings from cells)
- [x] Data validation (validate table data against rules)
- [x] Named range support (read named ranges as tables)
- [x] Cell comments support (read comments/notes from cells)
- [x] Hyperlink support (extract hyperlinks from cells)
- [x] Concurrent sheet processing (`ReadFileParallel()`)

### Medium Priority
- [ ] Streaming reader for large files
- [ ] Write support (create Excel files)

### Future
- [ ] gRPC/REST API wrapper
- [ ] Concurrent sheet processing
- [ ] Formula evaluation
- [ ] Legacy .xls format support

See [ROADMAP.md](ROADMAP.md) for detailed feature descriptions.

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Write tests for your changes
4. Ensure all tests pass (`go test ./...`)
5. Maintain or improve test coverage
6. Submit a pull request

Priority should be given to High Priority roadmap items.

## Acknowledgments

- [excelize](https://github.com/xuri/excelize) - Underlying Excel file parsing
- Inspired by the need for intelligent Excel parsing without predefined schemas
