# Goxls

A lightweight, high-performance Go library for reading Excel files (.xlsx) with automatic table detection and intelligent data extraction.

[![Go Version](https://img.shields.io/badge/Go-1.21+-00ADD8?style=flat&logo=go)](https://go.dev/)
[![CI](https://github.com/meddhiazoghlami/goxls/actions/workflows/ci.yml/badge.svg)](https://github.com/meddhiazoghlami/goxls/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/badge/coverage-97%25-brightgreen)](https://github.com/meddhiazoghlami/goxls)
[![Go Reference](https://pkg.go.dev/badge/github.com/meddhiazoghlami/goxls.svg)](https://pkg.go.dev/github.com/meddhiazoghlami/goxls)

## Overview

Goxls automatically detects and extracts tabular data from Excel spreadsheets without requiring predefined schemas or manual range specifications. It intelligently identifies table boundaries, headers, and data types, making it ideal for processing Excel files with unknown or dynamic structures.

## Features

| Category | Features |
|----------|----------|
| **Detection** | Automatic table detection, smart header detection, type inference |
| **Data Types** | Strings, numbers, dates, booleans, formulas |
| **Cell Features** | Merged cells, comments, hyperlinks, formulas |
| **Processing** | Multi-table support, concurrent sheet processing, named ranges |
| **Transformations** | Filter, Select, Rename, Reorder, Deduplicate |
| **Aggregations** | GroupBy, Sum, Count, Avg, Min, Max |
| **Validation** | Data validation rules, template validation |
| **Export** | JSON, CSV, SQL (MySQL, PostgreSQL, SQLite) |
| **Schema** | Go struct generation from tables |

## Installation

```bash
go get github.com/meddhiazoghlami/goxls
```

### Docker

```bash
# Pull or build
docker build -t goxls .

# Run
docker run --rm -v "$(pwd):/data" goxls myfile.xlsx
docker run --rm -v "$(pwd):/data" goxls -f json --pretty data.xlsx
```

## Quick Start

### Basic Usage

```go
package main

import (
    "fmt"
    "log"

    "github.com/meddhiazoghlami/goxls"
)

func main() {
    workbook, err := goxls.ReadFile("data.xlsx")
    if err != nil {
        log.Fatal(err)
    }

    for _, sheet := range workbook.Sheets {
        for _, table := range sheet.Tables {
            fmt.Printf("Table: %s (%d rows)\n", table.Name, table.RowCount())
            fmt.Printf("Headers: %v\n", table.Headers)
        }
    }
}
```

### With Options

```go
workbook, err := goxls.ReadFile("data.xlsx",
    goxls.WithMinColumns(3),
    goxls.WithMinRows(5),
    goxls.WithParallel(true),
    goxls.WithExpandMergedCells(true),
)
```

### With Context (Timeout/Cancellation)

```go
ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
defer cancel()

workbook, err := goxls.ReadFileWithContext(ctx, "large.xlsx")
```

### Error Handling

```go
workbook, err := goxls.ReadFile("data.xlsx")
if errors.Is(err, goxls.ErrFileNotFound) {
    log.Fatal("File does not exist")
}
if errors.Is(err, goxls.ErrNoTablesFound) {
    log.Fatal("No tables detected")
}
```

## Data Transformations

### Filtering

```go
// Filter rows with a predicate
adults := table.Filter(func(row goxls.Row) bool {
    if cell, ok := row.Get("Age"); ok {
        if val, ok := cell.AsFloat(); ok {
            return val >= 18
        }
    }
    return false
})

// Chain filters
result := table.Filter(isActive).Filter(hasEmail).Filter(isVerified)
```

### Column Operations

```go
// Select specific columns
subset := table.Select("Name", "Email", "Phone")

// Rename columns
renamed := table.Rename(map[string]string{
    "user_name": "Name",
    "user_email": "Email",
})

// Reorder columns
reordered := table.Reorder("Email", "Name", "Phone")

// Chain transformations
result := table.Select("name", "email").Rename(map[string]string{"name": "Name"})
```

### Deduplication

```go
// Find duplicate rows by key column
duplicates := table.FindDuplicates("Email")

// Remove duplicates (keep first occurrence)
unique := table.Deduplicate("Email")

// Get duplicate groups with counts
groups := table.FindDuplicateGroups("Email")
for _, g := range groups {
    fmt.Printf("Value: %s, Count: %d\n", g.KeyValue, g.Count)
}
```

## Aggregations

Perform SQL-like GROUP BY operations with aggregation functions.

```go
// Basic aggregation
result := table.GroupBy("Category").Aggregate(
    goxls.Sum("Amount"),
    goxls.Count("ID"),
    goxls.Avg("Price"),
    goxls.Min("Date"),
    goxls.Max("Date"),
)

// Multiple group columns
result := table.GroupBy("Region", "Category").Aggregate(
    goxls.Sum("Sales").As("TotalSales"),
    goxls.Count("OrderID").As("NumOrders"),
)

// Access results
for _, row := range result.Rows {
    category, _ := row.Get("Category")
    total, _ := row.Get("TotalSales")
    fmt.Printf("%s: %s\n", category.AsString(), total.AsString())
}
```

**Available Functions:**
| Function | Description |
|----------|-------------|
| `Sum(col)` | Sum of numeric values |
| `Count(col)` | Count of non-empty cells |
| `Avg(col)` | Average of numeric values |
| `Min(col)` | Minimum numeric value |
| `Max(col)` | Maximum numeric value |

Use `.As("alias")` to customize output column names.

## Column Analysis

```go
stats := table.AnalyzeColumns()

for _, col := range stats {
    fmt.Printf("Column: %s\n", col.Name)
    fmt.Printf("  Type: %v, Unique: %d, Empty: %d\n",
        col.InferredType, col.UniqueCount, col.EmptyCount)

    if col.HasNumericStats {
        fmt.Printf("  Min: %.2f, Max: %.2f, Avg: %.2f\n",
            col.Min, col.Max, col.Avg)
    }
}
```

## Table Comparison

```go
diff := goxls.DiffTables(oldTable, newTable, "ID")

if diff.HasChanges() {
    fmt.Printf("Added: %d, Removed: %d, Modified: %d\n",
        len(diff.AddedRows), len(diff.RemovedRows), len(diff.ModifiedRows))

    for _, mod := range diff.ModifiedRows {
        for _, change := range mod.Changes {
            fmt.Printf("%s: %s -> %s\n",
                change.Column, change.OldValue, change.NewValue)
        }
    }
}
```

## Export

### JSON

```go
jsonStr, _ := goxls.ToJSON(table)
jsonPretty, _ := goxls.ToJSONPretty(table)

// With options
opts := export.DefaultJSONOptions()
opts.Pretty = true
opts.SelectedColumns = []string{"Name", "Email"}
result, _ := export.NewJSONExporter(opts).ExportString(table)
```

### CSV

```go
csvStr, _ := goxls.ToCSV(table)
tsvStr, _ := goxls.ToTSV(table)
csvSemi, _ := goxls.ToCSVWithDelimiter(table, ';')
```

### SQL

```go
sqlStr, _ := goxls.ToSQL(table, "users")
sqlCreate, _ := goxls.ToSQLWithCreate(table, "users")

// With dialect
opts := export.DefaultSQLOptions()
opts.Dialect = export.DialectPostgreSQL
opts.CreateTable = true
opts.BatchSize = 100
result, _ := export.NewSQLExporter(opts).ExportString(table)
```

**Supported Dialects:** `DialectGeneric`, `DialectMySQL`, `DialectPostgreSQL`, `DialectSQLite`

## Validation

### Data Validation

```go
rules := []validation.ValidationRule{
    validation.ForColumn("Email").
        Required().
        MatchesPattern(`^[\w.-]+@[\w.-]+\.\w+$`).
        Build(),
    validation.ForColumn("Age").
        Range(18, 120).
        Build(),
    validation.ForColumn("Status").
        OneOf("active", "inactive", "pending").
        Build(),
}

result := validation.ValidateTable(table, rules)
if !result.Valid {
    for _, err := range result.Errors {
        fmt.Printf("Row %d, %s: %s\n", err.Row, err.Column, err.Message)
    }
}
```

### Template Validation

Validate workbook structure against a schema:

```go
template := goxls.NewTemplate("SalesTemplate").
    RequireSheets("Sales", "Inventory").
    StrictSheets().
    Sheet("Sales", goxls.NewSchema().
        RequireColumns("Date", "Amount", "Product").
        ColumnType("Amount", goxls.CellTypeNumber).
        RowCount(1, 1000).
        Build()).
    Build()

result := goxls.ValidateTemplate(workbook, template)
if !result.Valid {
    for _, err := range result.Errors {
        fmt.Printf("[%s] %s\n", err.Type.String(), err.Message)
    }
}
```

## Schema Generation

Generate Go structs from table headers:

```go
code, _ := goxls.GenerateStruct(table, "Person")
// Output:
// type Person struct {
//     Name   string  `excel:"Name"`
//     Age    float64 `excel:"Age"`
//     Email  string  `excel:"Email"`
// }

// With options
opts := &goxls.SchemaOptions{
    StructName:  "Employee",
    PackageName: "models",
    JSONTags:    true,
    OmitEmpty:   true,
}
code, _ := goxls.GenerateStructWithOptions(table, opts)
```

## Named Ranges

```go
nr := reader.NewNamedRangeReader()

// List named ranges
ranges, _ := nr.GetNamedRanges("data.xlsx")
for _, r := range ranges {
    fmt.Printf("%s -> %s\n", r.Name, r.RefersTo)
}

// Read as table
table, _ := nr.ReadRange("data.xlsx", "SalesData")
```

## Date Conversion

```go
import "github.com/meddhiazoghlami/goxls/pkg/dateutil"

// Excel serial to time.Time
t := dateutil.ExcelDateToTime(45658) // Jan 1, 2025

// time.Time to Excel serial
serial := dateutil.TimeToExcelDate(time.Now())

// Format directly
formatted := dateutil.FormatExcelDate(45658, "2006-01-02") // "2025-01-01"

// With timezone
loc, _ := time.LoadLocation("America/New_York")
t = dateutil.ExcelDateToTimeWithLocation(45658, loc)
```

## CLI

```bash
# Build
make build

# Basic usage
./bin/goxls data.xlsx

# Export formats
./bin/goxls -f json --pretty data.xlsx
./bin/goxls -f csv -o output.csv data.xlsx
./bin/goxls -f sql --sql-table=users data.xlsx

# Filtering
./bin/goxls -s Sales data.xlsx           # By sheet
./bin/goxls -t Sales_Table1 data.xlsx    # By table
./bin/goxls -c "Name,Email" data.xlsx    # By columns

# Summary
./bin/goxls --summary data.xlsx
```

| Option | Short | Description |
|--------|-------|-------------|
| `--format` | `-f` | Output: json, csv, sql, text |
| `--output` | `-o` | Output file (default: stdout) |
| `--sheet` | `-s` | Filter by sheet name |
| `--table` | `-t` | Filter by table name |
| `--columns` | `-c` | Columns to include |
| `--sql-table` | | SQL table name |
| `--summary` | | Show analysis summary |
| `--pretty` | | Pretty print JSON |

## Make Commands

```bash
make build        # Build CLI binary
make test         # Run tests
make cover        # Coverage report
make docker       # Build Docker image
make docker-test  # Test Docker image
make check        # Run fmt + vet + test
make help         # Show all commands
```

## Test Coverage

| Package | Coverage |
|---------|----------|
| `pkg/dateutil` | 100% |
| `pkg/models` | 99.0% |
| `pkg/validation` | 98.7% |
| `pkg/schema` | 96.6% |
| `pkg/reader` | 96.1% |
| `pkg/export` | 95.4% |
| **Total** | **97%** |

```bash
go test ./... -cover
```

## Limitations

- **Format:** Only .xlsx (Office 2007+), no .xls support
- **Memory:** Entire file loaded into memory; not suitable for very large files
- **Read-only:** Cannot create or modify Excel files
- **Formulas:** Extracted as strings, not evaluated

## Roadmap

See [ROADMAP.md](ROADMAP.md) for detailed feature plans.

**Completed:**
- Automatic table/header detection
- Type inference and column statistics
- Merged cells, comments, hyperlinks
- Data transformations (Filter, Select, Rename, Reorder)
- Aggregations (GroupBy, Sum, Count, Avg, Min, Max)
- Data and template validation
- Export (JSON, CSV, SQL)
- Schema generation
- Concurrent processing
- Docker support
- CLI

**Planned:**
- Streaming reader for large files
- Write support (create Excel files)
- Formula evaluation

## Acknowledgments

- [excelize](https://github.com/xuri/excelize) - Excel file parsing
