package export

import (
	"fmt"
	"io"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// Format represents the export format type
type Format int

const (
	FormatJSON Format = iota
	FormatCSV
	FormatSQL
)

// String returns the string representation of the format
func (f Format) String() string {
	switch f {
	case FormatJSON:
		return "json"
	case FormatCSV:
		return "csv"
	case FormatSQL:
		return "sql"
	default:
		return "unknown"
	}
}

// ParseFormat parses a string into a Format
func ParseFormat(s string) (Format, error) {
	switch s {
	case "json", "JSON":
		return FormatJSON, nil
	case "csv", "CSV":
		return FormatCSV, nil
	case "sql", "SQL":
		return FormatSQL, nil
	default:
		return 0, fmt.Errorf("unknown format: %s", s)
	}
}

// Exporter defines the interface for all exporters
type Exporter interface {
	// Export writes the table data to the writer
	Export(table *models.Table, w io.Writer) error

	// ExportBytes returns the exported data as bytes
	ExportBytes(table *models.Table) ([]byte, error)

	// ExportString returns the exported data as a string
	ExportString(table *models.Table) (string, error)
}

// Options holds common export options
type Options struct {
	// IncludeHeaders determines whether to include column headers
	IncludeHeaders bool

	// NullValue is the string to use for empty/null cells
	NullValue string

	// SelectedColumns limits export to specific columns (empty means all)
	SelectedColumns []string
}

// DefaultOptions returns sensible default options
func DefaultOptions() Options {
	return Options{
		IncludeHeaders:  true,
		NullValue:       "",
		SelectedColumns: nil,
	}
}

// Export exports a table to the specified format with default options
func Export(table *models.Table, format Format, w io.Writer) error {
	exporter, err := NewExporter(format, nil)
	if err != nil {
		return err
	}
	return exporter.Export(table, w)
}

// ExportString exports a table to a string in the specified format
func ExportString(table *models.Table, format Format) (string, error) {
	exporter, err := NewExporter(format, nil)
	if err != nil {
		return "", err
	}
	return exporter.ExportString(table)
}

// NewExporter creates an exporter for the given format
// Pass nil for opts to use defaults
func NewExporter(format Format, opts interface{}) (Exporter, error) {
	switch format {
	case FormatJSON:
		if opts == nil {
			return NewJSONExporter(nil), nil
		}
		if jsonOpts, ok := opts.(*JSONOptions); ok {
			return NewJSONExporter(jsonOpts), nil
		}
		return nil, fmt.Errorf("invalid options type for JSON exporter")

	case FormatCSV:
		if opts == nil {
			return NewCSVExporter(nil), nil
		}
		if csvOpts, ok := opts.(*CSVOptions); ok {
			return NewCSVExporter(csvOpts), nil
		}
		return nil, fmt.Errorf("invalid options type for CSV exporter")

	case FormatSQL:
		if opts == nil {
			return NewSQLExporter(nil), nil
		}
		if sqlOpts, ok := opts.(*SQLOptions); ok {
			return NewSQLExporter(sqlOpts), nil
		}
		return nil, fmt.Errorf("invalid options type for SQL exporter")

	default:
		return nil, fmt.Errorf("unsupported format: %v", format)
	}
}

// filterColumns returns headers and a filter function based on selected columns
func filterColumns(table *models.Table, selected []string) ([]string, map[string]bool) {
	if len(selected) == 0 {
		// Return all columns
		filter := make(map[string]bool, len(table.Headers))
		for _, h := range table.Headers {
			filter[h] = true
		}
		return table.Headers, filter
	}

	// Return only selected columns that exist
	filter := make(map[string]bool, len(selected))
	headers := make([]string, 0, len(selected))
	headerSet := make(map[string]bool, len(table.Headers))
	for _, h := range table.Headers {
		headerSet[h] = true
	}

	for _, col := range selected {
		if headerSet[col] {
			filter[col] = true
			headers = append(headers, col)
		}
	}
	return headers, filter
}

// getCellValue returns a normalized value for a cell
func getCellValue(cell models.Cell, nullValue string) interface{} {
	if cell.IsEmpty() {
		if nullValue != "" {
			return nullValue
		}
		return nil
	}
	return cell.Value
}
