// Package export provides functionality for exporting Excel table data to
// various formats including JSON, CSV, and SQL.
//
// # Supported Formats
//
//   - JSON: Array of objects with optional pretty-printing
//   - CSV: Comma-separated values with configurable delimiter
//   - SQL: INSERT statements with dialect support
//
// # Quick Export
//
// Simple one-line exports:
//
//	json, err := export.ToJSON(table)
//	jsonPretty, err := export.ToJSONPretty(table)
//	csv, err := export.ToCSV(table)
//	tsv, err := export.ToTSV(table)
//	sql, err := export.ToSQL(table, "users")
//
// # JSON Export
//
// Export with custom options:
//
//	opts := export.DefaultJSONOptions()
//	opts.Pretty = true
//	opts.SelectedColumns = []string{"Name", "Email"}
//
//	exporter := export.NewJSONExporter(opts)
//	result, err := exporter.ExportString(table)
//
// # CSV Export
//
// Export with custom delimiter:
//
//	opts := export.DefaultCSVOptions()
//	opts.Delimiter = ';'
//	opts.QuoteAll = true
//
//	exporter := export.NewCSVExporter(opts)
//	result, err := exporter.ExportString(table)
//
// # SQL Export
//
// Export with dialect support:
//
//	opts := export.DefaultSQLOptions()
//	opts.TableName = "employees"
//	opts.Dialect = export.DialectPostgreSQL
//	opts.CreateTable = true
//	opts.BatchSize = 100
//
//	exporter := export.NewSQLExporter(opts)
//	result, err := exporter.ExportString(table)
//
// # SQL Dialects
//
// Supported SQL dialects:
//
//   - DialectGeneric: Standard SQL
//   - DialectMySQL: MySQL-specific syntax
//   - DialectPostgreSQL: PostgreSQL-specific syntax
//   - DialectSQLite: SQLite-specific syntax
//
// # Writing to Files or Streams
//
// All exporters implement the Exporter interface:
//
//	type Exporter interface {
//	    Export(table *models.Table, w io.Writer) error
//	    ExportBytes(table *models.Table) ([]byte, error)
//	    ExportString(table *models.Table) (string, error)
//	}
package export
