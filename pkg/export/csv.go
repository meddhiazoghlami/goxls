package export

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"io"
	"time"

	"excel-lite/pkg/models"
)

// CSVOptions holds CSV-specific export options
type CSVOptions struct {
	Options

	// Delimiter is the field separator (default: ',')
	Delimiter rune

	// UseCRLF uses \r\n as line terminator (default: false, uses \n)
	UseCRLF bool

	// DateFormat is the format for date values (default: "2006-01-02")
	DateFormat string

	// QuoteAll forces quoting of all fields
	QuoteAll bool
}

// DefaultCSVOptions returns sensible defaults for CSV export
func DefaultCSVOptions() *CSVOptions {
	return &CSVOptions{
		Options:    DefaultOptions(),
		Delimiter:  ',',
		UseCRLF:    false,
		DateFormat: "2006-01-02",
		QuoteAll:   false,
	}
}

// CSVExporter exports tables to CSV format
type CSVExporter struct {
	opts *CSVOptions
}

// NewCSVExporter creates a new CSV exporter
func NewCSVExporter(opts *CSVOptions) *CSVExporter {
	if opts == nil {
		opts = DefaultCSVOptions()
	}
	return &CSVExporter{opts: opts}
}

// Export writes the table as CSV to the writer
func (e *CSVExporter) Export(table *models.Table, w io.Writer) error {
	headers, filter := filterColumns(table, e.opts.SelectedColumns)

	csvWriter := csv.NewWriter(w)
	csvWriter.Comma = e.opts.Delimiter
	csvWriter.UseCRLF = e.opts.UseCRLF

	// Write headers if enabled
	if e.opts.IncludeHeaders {
		if err := csvWriter.Write(headers); err != nil {
			return fmt.Errorf("failed to write headers: %w", err)
		}
	}

	// Write rows
	for _, row := range table.Rows {
		record := make([]string, 0, len(headers))
		for _, header := range headers {
			if filter[header] {
				cell, ok := row.Values[header]
				if ok {
					record = append(record, e.formatCell(cell))
				} else {
					record = append(record, e.opts.NullValue)
				}
			}
		}
		if err := csvWriter.Write(record); err != nil {
			return fmt.Errorf("failed to write row: %w", err)
		}
	}

	csvWriter.Flush()
	return csvWriter.Error()
}

// formatCell converts a cell to its string representation for CSV
func (e *CSVExporter) formatCell(cell models.Cell) string {
	if cell.IsEmpty() {
		return e.opts.NullValue
	}

	switch v := cell.Value.(type) {
	case time.Time:
		return v.Format(e.opts.DateFormat)
	case float64:
		// Use RawValue to preserve original formatting if available
		if cell.RawValue != "" {
			return cell.RawValue
		}
		return fmt.Sprintf("%g", v)
	case bool:
		if v {
			return "true"
		}
		return "false"
	case string:
		return v
	default:
		return cell.RawValue
	}
}

// ExportBytes returns the table as CSV bytes
func (e *CSVExporter) ExportBytes(table *models.Table) ([]byte, error) {
	buf := &bytes.Buffer{}
	if err := e.Export(table, buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// ExportString returns the table as a CSV string
func (e *CSVExporter) ExportString(table *models.Table) (string, error) {
	data, err := e.ExportBytes(table)
	if err != nil {
		return "", err
	}
	return string(data), nil
}

// ToCSV is a convenience method to export a table to CSV
func ToCSV(table *models.Table) (string, error) {
	return NewCSVExporter(nil).ExportString(table)
}

// ToCSVWithDelimiter exports a table to CSV with a custom delimiter
func ToCSVWithDelimiter(table *models.Table, delimiter rune) (string, error) {
	opts := DefaultCSVOptions()
	opts.Delimiter = delimiter
	return NewCSVExporter(opts).ExportString(table)
}

// ToTSV exports a table to tab-separated values
func ToTSV(table *models.Table) (string, error) {
	return ToCSVWithDelimiter(table, '\t')
}

// ToCSVWriter writes CSV to the provided writer
func ToCSVWriter(table *models.Table, w io.Writer) error {
	return NewCSVExporter(nil).Export(table, w)
}

// ToCSVBuffer writes CSV to a buffer and returns it
func ToCSVBuffer(table *models.Table) (*bytes.Buffer, error) {
	buf := &bytes.Buffer{}
	err := ToCSVWriter(table, buf)
	return buf, err
}
