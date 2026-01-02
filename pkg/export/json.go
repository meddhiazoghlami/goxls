package export

import (
	"bytes"
	"encoding/json"
	"io"

	"excel-lite/pkg/models"
)

// JSONOptions holds JSON-specific export options
type JSONOptions struct {
	Options

	// Pretty enables indented JSON output
	Pretty bool

	// Indent is the indentation string (default: "  ")
	Indent string

	// ArrayOnly outputs just the array without wrapping object
	ArrayOnly bool
}

// DefaultJSONOptions returns sensible defaults for JSON export
func DefaultJSONOptions() *JSONOptions {
	return &JSONOptions{
		Options:   DefaultOptions(),
		Pretty:    false,
		Indent:    "  ",
		ArrayOnly: false,
	}
}

// JSONExporter exports tables to JSON format
type JSONExporter struct {
	opts *JSONOptions
}

// NewJSONExporter creates a new JSON exporter
func NewJSONExporter(opts *JSONOptions) *JSONExporter {
	if opts == nil {
		opts = DefaultJSONOptions()
	}
	return &JSONExporter{opts: opts}
}

// Export writes the table as JSON to the writer
func (e *JSONExporter) Export(table *models.Table, w io.Writer) error {
	data, err := e.ExportBytes(table)
	if err != nil {
		return err
	}
	_, err = w.Write(data)
	return err
}

// ExportBytes returns the table as JSON bytes
func (e *JSONExporter) ExportBytes(table *models.Table) ([]byte, error) {
	headers, filter := filterColumns(table, e.opts.SelectedColumns)

	// Build rows as slice of maps
	rows := make([]map[string]interface{}, 0, len(table.Rows))
	for _, row := range table.Rows {
		rowMap := make(map[string]interface{}, len(headers))
		for _, header := range headers {
			if filter[header] {
				cell, ok := row.Values[header]
				if ok {
					rowMap[header] = getCellValue(cell, e.opts.NullValue)
				} else {
					rowMap[header] = nil
				}
			}
		}
		rows = append(rows, rowMap)
	}

	var output interface{}
	if e.opts.ArrayOnly {
		output = rows
	} else {
		output = map[string]interface{}{
			"name":    table.Name,
			"headers": headers,
			"rows":    rows,
			"count":   len(rows),
		}
	}

	if e.opts.Pretty {
		return json.MarshalIndent(output, "", e.opts.Indent)
	}
	return json.Marshal(output)
}

// ExportString returns the table as a JSON string
func (e *JSONExporter) ExportString(table *models.Table) (string, error) {
	data, err := e.ExportBytes(table)
	if err != nil {
		return "", err
	}
	return string(data), nil
}

// ToJSON is a convenience method to export a table to JSON
func ToJSON(table *models.Table) (string, error) {
	return NewJSONExporter(nil).ExportString(table)
}

// ToJSONPretty exports a table to pretty-printed JSON
func ToJSONPretty(table *models.Table) (string, error) {
	opts := DefaultJSONOptions()
	opts.Pretty = true
	return NewJSONExporter(opts).ExportString(table)
}

// ToJSONWriter writes JSON to the provided writer
func ToJSONWriter(table *models.Table, w io.Writer) error {
	return NewJSONExporter(nil).Export(table, w)
}

// ToJSONBuffer writes JSON to a buffer and returns it
func ToJSONBuffer(table *models.Table) (*bytes.Buffer, error) {
	buf := &bytes.Buffer{}
	err := ToJSONWriter(table, buf)
	return buf, err
}
