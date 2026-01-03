package export

import (
	"bytes"
	"fmt"
	"io"
	"regexp"
	"strings"
	"time"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// SQLDialect represents different SQL database dialects
type SQLDialect int

const (
	DialectGeneric SQLDialect = iota
	DialectMySQL
	DialectPostgreSQL
	DialectSQLite
)

// String returns the string representation of the dialect
func (d SQLDialect) String() string {
	switch d {
	case DialectMySQL:
		return "mysql"
	case DialectPostgreSQL:
		return "postgresql"
	case DialectSQLite:
		return "sqlite"
	default:
		return "generic"
	}
}

// SQLOptions holds SQL-specific export options
type SQLOptions struct {
	Options

	// TableName is the name of the SQL table (required)
	TableName string

	// Dialect specifies the SQL dialect for syntax variations
	Dialect SQLDialect

	// CreateTable includes CREATE TABLE statement
	CreateTable bool

	// DropTable includes DROP TABLE IF EXISTS before CREATE
	DropTable bool

	// BatchSize is the number of rows per INSERT statement (0 = all in one)
	BatchSize int

	// DateFormat is the format for date values
	DateFormat string
}

// DefaultSQLOptions returns sensible defaults for SQL export
func DefaultSQLOptions() *SQLOptions {
	return &SQLOptions{
		Options:     DefaultOptions(),
		TableName:   "exported_table",
		Dialect:     DialectGeneric,
		CreateTable: false,
		DropTable:   false,
		BatchSize:   0,
		DateFormat:  "2006-01-02 15:04:05",
	}
}

// SQLExporter exports tables to SQL INSERT statements
type SQLExporter struct {
	opts *SQLOptions
}

// NewSQLExporter creates a new SQL exporter
func NewSQLExporter(opts *SQLOptions) *SQLExporter {
	if opts == nil {
		opts = DefaultSQLOptions()
	}
	return &SQLExporter{opts: opts}
}

// Export writes the table as SQL to the writer
func (e *SQLExporter) Export(table *models.Table, w io.Writer) error {
	headers, filter := filterColumns(table, e.opts.SelectedColumns)

	// Write DROP TABLE if enabled
	if e.opts.DropTable {
		dropStmt := e.buildDropTable()
		if _, err := w.Write([]byte(dropStmt + "\n\n")); err != nil {
			return err
		}
	}

	// Write CREATE TABLE if enabled
	if e.opts.CreateTable {
		createStmt := e.buildCreateTable(table, headers)
		if _, err := w.Write([]byte(createStmt + "\n\n")); err != nil {
			return err
		}
	}

	// Write INSERT statements
	if len(table.Rows) == 0 {
		return nil
	}

	if e.opts.BatchSize <= 0 {
		// All rows in one INSERT
		insertStmt := e.buildInsert(table.Rows, headers, filter)
		if _, err := w.Write([]byte(insertStmt)); err != nil {
			return err
		}
	} else {
		// Batch inserts
		for i := 0; i < len(table.Rows); i += e.opts.BatchSize {
			end := i + e.opts.BatchSize
			if end > len(table.Rows) {
				end = len(table.Rows)
			}
			insertStmt := e.buildInsert(table.Rows[i:end], headers, filter)
			if _, err := w.Write([]byte(insertStmt)); err != nil {
				return err
			}
			if end < len(table.Rows) {
				if _, err := w.Write([]byte("\n")); err != nil {
					return err
				}
			}
		}
	}

	return nil
}

// buildDropTable generates a DROP TABLE statement
func (e *SQLExporter) buildDropTable() string {
	tableName := e.escapeIdentifier(e.opts.TableName)
	return fmt.Sprintf("DROP TABLE IF EXISTS %s;", tableName)
}

// buildCreateTable generates a CREATE TABLE statement
func (e *SQLExporter) buildCreateTable(table *models.Table, headers []string) string {
	tableName := e.escapeIdentifier(e.opts.TableName)

	var columns []string
	for _, header := range headers {
		colName := e.escapeIdentifier(header)
		colType := e.inferColumnType(table, header)
		columns = append(columns, fmt.Sprintf("    %s %s", colName, colType))
	}

	return fmt.Sprintf("CREATE TABLE %s (\n%s\n);", tableName, strings.Join(columns, ",\n"))
}

// inferColumnType attempts to infer SQL column type from table data
func (e *SQLExporter) inferColumnType(table *models.Table, header string) string {
	var hasString, hasNumber, hasDate, hasBool bool

	for _, row := range table.Rows {
		cell, ok := row.Values[header]
		if !ok || cell.IsEmpty() {
			continue
		}

		switch cell.Type {
		case models.CellTypeString:
			hasString = true
		case models.CellTypeNumber:
			hasNumber = true
		case models.CellTypeDate:
			hasDate = true
		case models.CellTypeBool:
			hasBool = true
		}
	}

	// Determine type based on what we found
	switch {
	case hasString:
		return e.stringType()
	case hasDate:
		return e.dateType()
	case hasBool && !hasNumber:
		return e.boolType()
	case hasNumber:
		return e.numberType()
	default:
		return e.stringType()
	}
}

// Type helpers for different dialects
func (e *SQLExporter) stringType() string {
	switch e.opts.Dialect {
	case DialectPostgreSQL:
		return "TEXT"
	case DialectMySQL:
		return "VARCHAR(255)"
	default:
		return "TEXT"
	}
}

func (e *SQLExporter) numberType() string {
	switch e.opts.Dialect {
	case DialectPostgreSQL:
		return "NUMERIC"
	case DialectMySQL:
		return "DOUBLE"
	default:
		return "REAL"
	}
}

func (e *SQLExporter) dateType() string {
	switch e.opts.Dialect {
	case DialectPostgreSQL:
		return "TIMESTAMP"
	case DialectMySQL:
		return "DATETIME"
	default:
		return "TEXT"
	}
}

func (e *SQLExporter) boolType() string {
	switch e.opts.Dialect {
	case DialectPostgreSQL:
		return "BOOLEAN"
	case DialectMySQL:
		return "TINYINT(1)"
	default:
		return "INTEGER"
	}
}

// buildInsert generates an INSERT statement for the given rows
func (e *SQLExporter) buildInsert(rows []models.Row, headers []string, filter map[string]bool) string {
	tableName := e.escapeIdentifier(e.opts.TableName)

	// Build column list
	var escapedHeaders []string
	for _, h := range headers {
		escapedHeaders = append(escapedHeaders, e.escapeIdentifier(h))
	}
	columnList := strings.Join(escapedHeaders, ", ")

	// Build values
	var valueGroups []string
	for _, row := range rows {
		var values []string
		for _, header := range headers {
			if filter[header] {
				cell, ok := row.Values[header]
				if ok {
					values = append(values, e.formatValue(cell))
				} else {
					values = append(values, "NULL")
				}
			}
		}
		valueGroups = append(valueGroups, "("+strings.Join(values, ", ")+")")
	}

	return fmt.Sprintf("INSERT INTO %s (%s) VALUES\n%s;",
		tableName, columnList, strings.Join(valueGroups, ",\n"))
}

// escapeIdentifier escapes a SQL identifier (table/column name)
func (e *SQLExporter) escapeIdentifier(name string) string {
	// Remove any existing quotes and dangerous characters
	clean := regexp.MustCompile(`[^\w]`).ReplaceAllString(name, "_")

	switch e.opts.Dialect {
	case DialectMySQL:
		return "`" + clean + "`"
	case DialectPostgreSQL:
		return `"` + clean + `"`
	default:
		return `"` + clean + `"`
	}
}

// formatValue formats a cell value for SQL
func (e *SQLExporter) formatValue(cell models.Cell) string {
	if cell.IsEmpty() {
		return "NULL"
	}

	switch v := cell.Value.(type) {
	case time.Time:
		return fmt.Sprintf("'%s'", v.Format(e.opts.DateFormat))
	case float64:
		return fmt.Sprintf("%g", v)
	case bool:
		switch e.opts.Dialect {
		case DialectPostgreSQL:
			if v {
				return "TRUE"
			}
			return "FALSE"
		default:
			if v {
				return "1"
			}
			return "0"
		}
	case string:
		return e.escapeString(v)
	default:
		return e.escapeString(cell.RawValue)
	}
}

// escapeString escapes a string value for SQL
func (e *SQLExporter) escapeString(s string) string {
	// Escape single quotes by doubling them
	escaped := strings.ReplaceAll(s, "'", "''")
	return "'" + escaped + "'"
}

// ExportBytes returns the table as SQL bytes
func (e *SQLExporter) ExportBytes(table *models.Table) ([]byte, error) {
	buf := &bytes.Buffer{}
	if err := e.Export(table, buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// ExportString returns the table as SQL string
func (e *SQLExporter) ExportString(table *models.Table) (string, error) {
	data, err := e.ExportBytes(table)
	if err != nil {
		return "", err
	}
	return string(data), nil
}

// ToSQL is a convenience method to export a table to SQL
func ToSQL(table *models.Table, tableName string) (string, error) {
	opts := DefaultSQLOptions()
	opts.TableName = tableName
	return NewSQLExporter(opts).ExportString(table)
}

// ToSQLWithCreate exports a table to SQL with CREATE TABLE statement
func ToSQLWithCreate(table *models.Table, tableName string) (string, error) {
	opts := DefaultSQLOptions()
	opts.TableName = tableName
	opts.CreateTable = true
	return NewSQLExporter(opts).ExportString(table)
}

// ToSQLWriter writes SQL to the provided writer
func ToSQLWriter(table *models.Table, tableName string, w io.Writer) error {
	opts := DefaultSQLOptions()
	opts.TableName = tableName
	return NewSQLExporter(opts).Export(table, w)
}

// ToSQLBuffer writes SQL to a buffer and returns it
func ToSQLBuffer(table *models.Table, tableName string) (*bytes.Buffer, error) {
	buf := &bytes.Buffer{}
	err := ToSQLWriter(table, tableName, buf)
	return buf, err
}
