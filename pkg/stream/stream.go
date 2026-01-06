// Package stream provides streaming functionality for reading large Excel files
// row-by-row without loading the entire file into memory.
//
// Basic usage:
//
//	sr, err := stream.NewStreamReader("huge_file.xlsx", "Sheet1")
//	if err != nil {
//	    log.Fatal(err)
//	}
//	defer sr.Close()
//
//	for {
//	    row, err := sr.Next()
//	    if err == io.EOF {
//	        break
//	    }
//	    if err != nil {
//	        log.Fatal(err)
//	    }
//	    // Process row
//	}
package stream

import (
	"context"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/meddhiazoghlami/goxls/pkg/models"
	"github.com/xuri/excelize/v2"
)

// Sentinel errors for streaming operations
var (
	// ErrStreamClosed is returned when operations are attempted on a closed stream
	ErrStreamClosed = errors.New("goxls: stream reader is closed")

	// ErrNoHeaders is returned when headers are required but not available
	ErrNoHeaders = errors.New("goxls: no headers available for stream")

	// ErrInvalidStreamConfig is returned for invalid streaming configuration
	ErrInvalidStreamConfig = errors.New("goxls: invalid stream configuration")
)

// StreamConfig holds configuration for streaming operations
type StreamConfig struct {
	// HasHeaders indicates whether the first data row contains headers (default: true)
	HasHeaders bool

	// Headers provides explicit column headers, overriding auto-detection
	Headers []string

	// SkipRows is the number of rows to skip before reading headers/data
	SkipRows int

	// DetectTypes enables type inference from raw string values (default: true)
	DetectTypes bool

	// DateFormats specifies custom date formats for type detection
	DateFormats []string

	// TrimSpaces trims whitespace from cell values (default: true)
	TrimSpaces bool

	// SkipEmptyRows skips rows where all cells are empty (default: true)
	SkipEmptyRows bool
}

// DefaultStreamConfig returns the default streaming configuration
func DefaultStreamConfig() StreamConfig {
	return StreamConfig{
		HasHeaders:    true,
		Headers:       nil,
		SkipRows:      0,
		DetectTypes:   true,
		DateFormats:   nil,
		TrimSpaces:    true,
		SkipEmptyRows: true,
	}
}

// StreamCell represents a single cell from streaming with inferred type
type StreamCell struct {
	// Value is the parsed value (string, float64, bool, time.Time, or nil)
	Value interface{}

	// Type is the inferred cell type
	Type models.CellType

	// RawValue is the original string value from Excel
	RawValue string

	// ColIndex is the 0-indexed column position
	ColIndex int
}

// IsEmpty returns true if the cell is empty
func (c *StreamCell) IsEmpty() bool {
	return c.Type == models.CellTypeEmpty || c.RawValue == ""
}

// AsString returns the cell value as a string
func (c *StreamCell) AsString() string {
	if c.Value == nil {
		return ""
	}
	if s, ok := c.Value.(string); ok {
		return s
	}
	return c.RawValue
}

// AsFloat returns the cell value as a float64 if it's numeric
func (c *StreamCell) AsFloat() (float64, bool) {
	if v, ok := c.Value.(float64); ok {
		return v, true
	}
	return 0, false
}

// StreamRow represents a single row from streaming
type StreamRow struct {
	// Index is the 0-indexed row number in the file (after skipped rows)
	Index int

	// Values maps column headers to cells
	Values map[string]StreamCell

	// Cells contains all cells in column order
	Cells []StreamCell
}

// Get returns the cell for a given header name
func (r *StreamRow) Get(header string) (StreamCell, bool) {
	cell, ok := r.Values[header]
	return cell, ok
}

// IsEmpty returns true if all cells in the row are empty
func (r *StreamRow) IsEmpty() bool {
	for _, cell := range r.Cells {
		if !cell.IsEmpty() {
			return false
		}
	}
	return true
}

// StreamReader provides row-by-row iteration over Excel sheet data
type StreamReader struct {
	file       *excelize.File
	rows       *excelize.Rows
	sheetName  string
	filePath   string
	headers    []string
	currentRow int
	dataRowNum int // tracks actual data rows read (excludes headers/skipped)
	config     StreamConfig
	typeInfer  *TypeInferrer
	ctx        context.Context
	closed     bool
	err        error
}

// NewStreamReader creates a new streaming reader for the specified sheet.
// It opens the file and prepares for row-by-row iteration.
//
// Example:
//
//	sr, err := stream.NewStreamReader("data.xlsx", "Sheet1")
//	if err != nil {
//	    log.Fatal(err)
//	}
//	defer sr.Close()
func NewStreamReader(filePath, sheetName string, opts ...StreamOption) (*StreamReader, error) {
	return NewStreamReaderWithContext(context.Background(), filePath, sheetName, opts...)
}

// NewStreamReaderWithContext creates a streaming reader with context support for cancellation.
//
// Example:
//
//	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Minute)
//	defer cancel()
//	sr, err := stream.NewStreamReaderWithContext(ctx, "huge.xlsx", "Sheet1")
func NewStreamReaderWithContext(ctx context.Context, filePath, sheetName string, opts ...StreamOption) (*StreamReader, error) {
	// Check context before starting
	select {
	case <-ctx.Done():
		return nil, fmt.Errorf("goxls: operation canceled: %w", ctx.Err())
	default:
	}

	// Validate file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, fmt.Errorf("goxls: file not found: %s", filePath)
	}

	// Apply options
	o := &streamOptions{
		config: DefaultStreamConfig(),
	}
	for _, opt := range opts {
		opt(o)
	}

	// Open the Excel file
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("goxls: failed to open file: %w", err)
	}

	// Verify sheet exists
	sheetIndex, err := f.GetSheetIndex(sheetName)
	if err != nil || sheetIndex == -1 {
		f.Close()
		return nil, fmt.Errorf("goxls: sheet not found: %s", sheetName)
	}

	// Create row iterator
	rows, err := f.Rows(sheetName)
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("goxls: failed to create row iterator: %w", err)
	}

	sr := &StreamReader{
		file:       f,
		rows:       rows,
		sheetName:  sheetName,
		filePath:   filePath,
		config:     o.config,
		typeInfer:  NewTypeInferrer(o.config.DateFormats),
		ctx:        ctx,
		currentRow: 0,
		dataRowNum: 0,
	}

	// Skip initial rows if configured
	for i := 0; i < o.config.SkipRows; i++ {
		if !rows.Next() {
			break
		}
		sr.currentRow++
	}

	// Read headers
	if err := sr.initHeaders(); err != nil {
		sr.Close()
		return nil, err
	}

	return sr, nil
}

// initHeaders reads or sets up headers based on configuration
func (sr *StreamReader) initHeaders() error {
	// If explicit headers provided, use them
	if len(sr.config.Headers) > 0 {
		sr.headers = make([]string, len(sr.config.Headers))
		copy(sr.headers, sr.config.Headers)
		return nil
	}

	// If no headers expected, generate column names on first row
	if !sr.config.HasHeaders {
		// Headers will be generated when we see the first row
		return nil
	}

	// Read header row
	if !sr.rows.Next() {
		if err := sr.rows.Error(); err != nil {
			return fmt.Errorf("goxls: failed to read header row: %w", err)
		}
		return ErrNoHeaders
	}
	sr.currentRow++

	cols, err := sr.rows.Columns()
	if err != nil {
		return fmt.Errorf("goxls: failed to read header columns: %w", err)
	}

	sr.headers = sr.normalizeHeaders(cols)
	return nil
}

// normalizeHeaders cleans up header names
func (sr *StreamReader) normalizeHeaders(raw []string) []string {
	headers := make([]string, len(raw))
	seen := make(map[string]int)

	for i, h := range raw {
		// Trim whitespace if configured
		if sr.config.TrimSpaces {
			h = strings.TrimSpace(h)
		}

		// Handle empty headers
		if h == "" {
			h = fmt.Sprintf("Column_%d", i+1)
		}

		// Handle duplicates
		if count, exists := seen[h]; exists {
			seen[h] = count + 1
			h = fmt.Sprintf("%s_%d", h, count+1)
		} else {
			seen[h] = 1
		}

		headers[i] = h
	}

	return headers
}

// generateHeaders creates default column names for no-header mode
func (sr *StreamReader) generateHeaders(colCount int) {
	sr.headers = make([]string, colCount)
	for i := 0; i < colCount; i++ {
		sr.headers[i] = fmt.Sprintf("Column_%d", i+1)
	}
}

// Next advances to the next row and returns it.
// Returns io.EOF when no more rows are available.
// Other errors indicate read failures.
func (sr *StreamReader) Next() (*StreamRow, error) {
	if sr.closed {
		return nil, ErrStreamClosed
	}

	if sr.err != nil {
		return nil, sr.err
	}

	// Check context
	select {
	case <-sr.ctx.Done():
		sr.err = fmt.Errorf("goxls: operation canceled: %w", sr.ctx.Err())
		return nil, sr.err
	default:
	}

	for {
		if !sr.rows.Next() {
			// Check for iteration error
			if err := sr.rows.Error(); err != nil {
				sr.err = fmt.Errorf("goxls: stream read error: %w", err)
				return nil, sr.err
			}
			return nil, io.EOF
		}
		sr.currentRow++

		cols, err := sr.rows.Columns()
		if err != nil {
			sr.err = fmt.Errorf("goxls: failed to read columns: %w", err)
			return nil, sr.err
		}

		// Generate headers if in no-header mode and this is the first data row
		if sr.headers == nil && !sr.config.HasHeaders {
			sr.generateHeaders(len(cols))
		}

		// Build the row
		row := sr.buildRow(cols)

		// Skip empty rows if configured
		if sr.config.SkipEmptyRows && row.IsEmpty() {
			continue
		}

		sr.dataRowNum++
		return row, nil
	}
}

// buildRow constructs a StreamRow from raw column values
func (sr *StreamReader) buildRow(cols []string) *StreamRow {
	row := &StreamRow{
		Index:  sr.dataRowNum,
		Values: make(map[string]StreamCell),
		Cells:  make([]StreamCell, 0, len(cols)),
	}

	// Ensure we have enough headers
	maxCols := len(cols)
	if len(sr.headers) > maxCols {
		maxCols = len(sr.headers)
	}

	for i := 0; i < maxCols; i++ {
		var rawValue string
		if i < len(cols) {
			rawValue = cols[i]
		}

		// Trim whitespace if configured
		if sr.config.TrimSpaces {
			rawValue = strings.TrimSpace(rawValue)
		}

		// Detect type and parse value
		var cellType models.CellType
		var value interface{}

		if sr.config.DetectTypes {
			cellType = sr.typeInfer.InferType(rawValue)
			value = sr.typeInfer.ParseValue(rawValue, cellType)
		} else {
			if rawValue == "" {
				cellType = models.CellTypeEmpty
				value = nil
			} else {
				cellType = models.CellTypeString
				value = rawValue
			}
		}

		cell := StreamCell{
			Value:    value,
			Type:     cellType,
			RawValue: rawValue,
			ColIndex: i,
		}

		row.Cells = append(row.Cells, cell)

		// Map to header if available
		if i < len(sr.headers) {
			row.Values[sr.headers[i]] = cell
		}
	}

	return row
}

// Headers returns the column headers.
// These are either auto-detected from the first row, explicitly provided,
// or generated as Column_1, Column_2, etc.
func (sr *StreamReader) Headers() []string {
	if sr.headers == nil {
		return []string{}
	}
	result := make([]string, len(sr.headers))
	copy(result, sr.headers)
	return result
}

// CurrentRow returns the current row index in the file (0-indexed, includes skipped rows)
func (sr *StreamReader) CurrentRow() int {
	return sr.currentRow
}

// TotalRowsRead returns the count of data rows read so far (excludes headers and skipped rows)
func (sr *StreamReader) TotalRowsRead() int {
	return sr.dataRowNum
}

// SheetName returns the name of the sheet being read
func (sr *StreamReader) SheetName() string {
	return sr.sheetName
}

// FilePath returns the path of the file being read
func (sr *StreamReader) FilePath() string {
	return sr.filePath
}

// Close releases all resources associated with the stream reader.
// Must be called when done reading to prevent resource leaks.
func (sr *StreamReader) Close() error {
	if sr.closed {
		return nil
	}
	sr.closed = true

	var errs []error

	if sr.rows != nil {
		if err := sr.rows.Close(); err != nil {
			errs = append(errs, err)
		}
	}

	if sr.file != nil {
		if err := sr.file.Close(); err != nil {
			errs = append(errs, err)
		}
	}

	if len(errs) > 0 {
		return fmt.Errorf("goxls: errors closing stream: %v", errs)
	}
	return nil
}

// ForEach iterates over all rows and calls the provided function for each.
// Iteration stops if the function returns an error.
// Returns nil when all rows have been processed successfully.
//
// Example:
//
//	err := sr.ForEach(func(row *StreamRow) error {
//	    fmt.Println(row.Values["Name"])
//	    return nil
//	})
func (sr *StreamReader) ForEach(fn func(*StreamRow) error) error {
	for {
		row, err := sr.Next()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}

		if err := fn(row); err != nil {
			return err
		}
	}
}

// CollectN reads up to n rows and returns them as a slice.
// Returns fewer rows if EOF is reached before n rows.
// Warning: For large n, this defeats the purpose of streaming.
func (sr *StreamReader) CollectN(n int) ([]*StreamRow, error) {
	rows := make([]*StreamRow, 0, n)
	for i := 0; i < n; i++ {
		row, err := sr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return rows, err
		}
		rows = append(rows, row)
	}
	return rows, nil
}

// Collect reads all remaining rows and returns them as a slice.
// Warning: This loads all remaining data into memory, defeating the purpose of streaming.
// Use only for small datasets or when full collection is necessary.
func (sr *StreamReader) Collect() ([]*StreamRow, error) {
	var rows []*StreamRow
	for {
		row, err := sr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return rows, err
		}
		rows = append(rows, row)
	}
	return rows, nil
}
