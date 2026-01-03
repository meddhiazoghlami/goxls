package export

import (
	"bytes"
	"encoding/json"
	"fmt"
	"strings"
	"testing"
	"time"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// Helper function to create a test table
func createTestTable() *models.Table {
	return &models.Table{
		Name:    "TestTable",
		Headers: []string{"ID", "Name", "Age", "Active", "JoinDate"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"ID":       {Value: float64(1), Type: models.CellTypeNumber, RawValue: "1"},
					"Name":     {Value: "Alice", Type: models.CellTypeString, RawValue: "Alice"},
					"Age":      {Value: float64(30), Type: models.CellTypeNumber, RawValue: "30"},
					"Active":   {Value: true, Type: models.CellTypeBool, RawValue: "true"},
					"JoinDate": {Value: time.Date(2023, 1, 15, 0, 0, 0, 0, time.UTC), Type: models.CellTypeDate, RawValue: "2023-01-15"},
				},
			},
			{
				Index: 2,
				Values: map[string]models.Cell{
					"ID":       {Value: float64(2), Type: models.CellTypeNumber, RawValue: "2"},
					"Name":     {Value: "Bob", Type: models.CellTypeString, RawValue: "Bob"},
					"Age":      {Value: float64(25), Type: models.CellTypeNumber, RawValue: "25"},
					"Active":   {Value: false, Type: models.CellTypeBool, RawValue: "false"},
					"JoinDate": {Value: time.Date(2023, 6, 20, 0, 0, 0, 0, time.UTC), Type: models.CellTypeDate, RawValue: "2023-06-20"},
				},
			},
			{
				Index: 3,
				Values: map[string]models.Cell{
					"ID":       {Value: float64(3), Type: models.CellTypeNumber, RawValue: "3"},
					"Name":     {Value: "Charlie", Type: models.CellTypeString, RawValue: "Charlie"},
					"Age":      {Value: nil, Type: models.CellTypeEmpty, RawValue: ""},
					"Active":   {Value: true, Type: models.CellTypeBool, RawValue: "true"},
					"JoinDate": {Value: nil, Type: models.CellTypeEmpty, RawValue: ""},
				},
			},
		},
	}
}

func createEmptyTable() *models.Table {
	return &models.Table{
		Name:    "EmptyTable",
		Headers: []string{"A", "B"},
		Rows:    []models.Row{},
	}
}

// ============ Format Tests ============

func TestFormatString(t *testing.T) {
	tests := []struct {
		format   Format
		expected string
	}{
		{FormatJSON, "json"},
		{FormatCSV, "csv"},
		{FormatSQL, "sql"},
		{Format(99), "unknown"},
	}

	for _, tt := range tests {
		if got := tt.format.String(); got != tt.expected {
			t.Errorf("Format(%d).String() = %q, want %q", tt.format, got, tt.expected)
		}
	}
}

func TestParseFormat(t *testing.T) {
	tests := []struct {
		input    string
		expected Format
		wantErr  bool
	}{
		{"json", FormatJSON, false},
		{"JSON", FormatJSON, false},
		{"csv", FormatCSV, false},
		{"CSV", FormatCSV, false},
		{"sql", FormatSQL, false},
		{"SQL", FormatSQL, false},
		{"xml", 0, true},
		{"", 0, true},
	}

	for _, tt := range tests {
		got, err := ParseFormat(tt.input)
		if (err != nil) != tt.wantErr {
			t.Errorf("ParseFormat(%q) error = %v, wantErr %v", tt.input, err, tt.wantErr)
			continue
		}
		if !tt.wantErr && got != tt.expected {
			t.Errorf("ParseFormat(%q) = %v, want %v", tt.input, got, tt.expected)
		}
	}
}

func TestNewExporter(t *testing.T) {
	tests := []struct {
		name    string
		format  Format
		opts    interface{}
		wantErr bool
	}{
		{"JSON default", FormatJSON, nil, false},
		{"JSON with opts", FormatJSON, &JSONOptions{}, false},
		{"JSON wrong opts", FormatJSON, &CSVOptions{}, true},
		{"CSV default", FormatCSV, nil, false},
		{"CSV with opts", FormatCSV, &CSVOptions{}, false},
		{"CSV wrong opts", FormatCSV, &SQLOptions{}, true},
		{"SQL default", FormatSQL, nil, false},
		{"SQL with opts", FormatSQL, &SQLOptions{}, false},
		{"SQL wrong opts", FormatSQL, &JSONOptions{}, true},
		{"Unknown format", Format(99), nil, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := NewExporter(tt.format, tt.opts)
			if (err != nil) != tt.wantErr {
				t.Errorf("NewExporter() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestExportConvenienceFunction(t *testing.T) {
	table := createTestTable()
	buf := &bytes.Buffer{}

	err := Export(table, FormatJSON, buf)
	if err != nil {
		t.Fatalf("Export() error = %v", err)
	}

	if buf.Len() == 0 {
		t.Error("Export() produced empty output")
	}
}

func TestExportStringConvenienceFunction(t *testing.T) {
	table := createTestTable()

	result, err := ExportString(table, FormatCSV)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	if result == "" {
		t.Error("ExportString() produced empty output")
	}
}

// ============ JSON Tests ============

func TestJSONExporter(t *testing.T) {
	table := createTestTable()
	exporter := NewJSONExporter(nil)

	result, err := exporter.ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	// Verify it's valid JSON
	var data map[string]interface{}
	if err := json.Unmarshal([]byte(result), &data); err != nil {
		t.Fatalf("Result is not valid JSON: %v", err)
	}

	// Check structure
	if data["name"] != "TestTable" {
		t.Errorf("Expected name 'TestTable', got %v", data["name"])
	}

	if int(data["count"].(float64)) != 3 {
		t.Errorf("Expected count 3, got %v", data["count"])
	}

	rows, ok := data["rows"].([]interface{})
	if !ok || len(rows) != 3 {
		t.Errorf("Expected 3 rows, got %v", data["rows"])
	}
}

func TestJSONExporterPretty(t *testing.T) {
	table := createTestTable()
	opts := DefaultJSONOptions()
	opts.Pretty = true

	result, err := NewJSONExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	// Pretty JSON should have newlines
	if !strings.Contains(result, "\n") {
		t.Error("Pretty JSON should contain newlines")
	}
}

func TestJSONExporterArrayOnly(t *testing.T) {
	table := createTestTable()
	opts := DefaultJSONOptions()
	opts.ArrayOnly = true

	result, err := NewJSONExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	// Should be an array, not an object
	if !strings.HasPrefix(result, "[") {
		t.Error("ArrayOnly output should start with '['")
	}
}

func TestJSONExporterSelectedColumns(t *testing.T) {
	table := createTestTable()
	opts := DefaultJSONOptions()
	opts.SelectedColumns = []string{"Name", "Age"}

	result, err := NewJSONExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	var data map[string]interface{}
	json.Unmarshal([]byte(result), &data)

	headers := data["headers"].([]interface{})
	if len(headers) != 2 {
		t.Errorf("Expected 2 headers, got %d", len(headers))
	}
}

func TestJSONConvenienceFunctions(t *testing.T) {
	table := createTestTable()

	// ToJSON
	result, err := ToJSON(table)
	if err != nil {
		t.Errorf("ToJSON() error = %v", err)
	}
	if result == "" {
		t.Error("ToJSON() returned empty string")
	}

	// ToJSONPretty
	result, err = ToJSONPretty(table)
	if err != nil {
		t.Errorf("ToJSONPretty() error = %v", err)
	}
	if !strings.Contains(result, "\n") {
		t.Error("ToJSONPretty() should have newlines")
	}

	// ToJSONBuffer
	buf, err := ToJSONBuffer(table)
	if err != nil {
		t.Errorf("ToJSONBuffer() error = %v", err)
	}
	if buf.Len() == 0 {
		t.Error("ToJSONBuffer() returned empty buffer")
	}
}

// ============ CSV Tests ============

func TestCSVExporter(t *testing.T) {
	table := createTestTable()
	exporter := NewCSVExporter(nil)

	result, err := exporter.ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	lines := strings.Split(strings.TrimSpace(result), "\n")
	if len(lines) != 4 { // header + 3 rows
		t.Errorf("Expected 4 lines, got %d", len(lines))
	}

	// Check header
	if !strings.Contains(lines[0], "ID") || !strings.Contains(lines[0], "Name") {
		t.Error("Header line should contain column names")
	}
}

func TestCSVExporterNoHeaders(t *testing.T) {
	table := createTestTable()
	opts := DefaultCSVOptions()
	opts.IncludeHeaders = false

	result, err := NewCSVExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	lines := strings.Split(strings.TrimSpace(result), "\n")
	if len(lines) != 3 { // 3 rows, no header
		t.Errorf("Expected 3 lines, got %d", len(lines))
	}
}

func TestCSVExporterCustomDelimiter(t *testing.T) {
	table := createTestTable()
	opts := DefaultCSVOptions()
	opts.Delimiter = ';'

	result, err := NewCSVExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	if !strings.Contains(result, ";") {
		t.Error("Result should contain semicolons")
	}
}

func TestCSVExporterDateFormat(t *testing.T) {
	table := createTestTable()
	opts := DefaultCSVOptions()
	opts.DateFormat = "02/01/2006"

	result, err := NewCSVExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	if !strings.Contains(result, "15/01/2023") {
		t.Errorf("Date should be formatted as DD/MM/YYYY, got: %s", result)
	}
}

func TestCSVConvenienceFunctions(t *testing.T) {
	table := createTestTable()

	// ToCSV
	result, err := ToCSV(table)
	if err != nil {
		t.Errorf("ToCSV() error = %v", err)
	}
	if result == "" {
		t.Error("ToCSV() returned empty string")
	}

	// ToTSV
	result, err = ToTSV(table)
	if err != nil {
		t.Errorf("ToTSV() error = %v", err)
	}
	if !strings.Contains(result, "\t") {
		t.Error("ToTSV() should contain tabs")
	}

	// ToCSVBuffer
	buf, err := ToCSVBuffer(table)
	if err != nil {
		t.Errorf("ToCSVBuffer() error = %v", err)
	}
	if buf.Len() == 0 {
		t.Error("ToCSVBuffer() returned empty buffer")
	}
}

// ============ SQL Tests ============

func TestSQLExporter(t *testing.T) {
	table := createTestTable()
	opts := DefaultSQLOptions()
	opts.TableName = "users"

	result, err := NewSQLExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	if !strings.Contains(result, "INSERT INTO") {
		t.Error("Result should contain INSERT INTO")
	}
	if !strings.Contains(result, `"users"`) {
		t.Errorf("Result should contain table name, got: %s", result)
	}
}

func TestSQLExporterWithCreate(t *testing.T) {
	table := createTestTable()
	opts := DefaultSQLOptions()
	opts.TableName = "users"
	opts.CreateTable = true

	result, err := NewSQLExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	if !strings.Contains(result, "CREATE TABLE") {
		t.Error("Result should contain CREATE TABLE")
	}
}

func TestSQLExporterWithDrop(t *testing.T) {
	table := createTestTable()
	opts := DefaultSQLOptions()
	opts.TableName = "users"
	opts.DropTable = true

	result, err := NewSQLExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	if !strings.Contains(result, "DROP TABLE IF EXISTS") {
		t.Error("Result should contain DROP TABLE IF EXISTS")
	}
}

func TestSQLExporterBatchSize(t *testing.T) {
	table := createTestTable()
	opts := DefaultSQLOptions()
	opts.TableName = "users"
	opts.BatchSize = 2

	result, err := NewSQLExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	// Should have 2 INSERT statements
	insertCount := strings.Count(result, "INSERT INTO")
	if insertCount != 2 {
		t.Errorf("Expected 2 INSERT statements, got %d", insertCount)
	}
}

func TestSQLExporterDialects(t *testing.T) {
	table := createTestTable()

	tests := []struct {
		dialect      SQLDialect
		expectQuote  string
		expectBool   string
	}{
		{DialectMySQL, "`", "1"},
		{DialectPostgreSQL, `"`, "TRUE"},
		{DialectSQLite, `"`, "1"},
		{DialectGeneric, `"`, "1"},
	}

	for _, tt := range tests {
		t.Run(tt.dialect.String(), func(t *testing.T) {
			opts := DefaultSQLOptions()
			opts.TableName = "users"
			opts.Dialect = tt.dialect

			result, err := NewSQLExporter(opts).ExportString(table)
			if err != nil {
				t.Fatalf("ExportString() error = %v", err)
			}

			if !strings.Contains(result, tt.expectQuote+"users"+tt.expectQuote) {
				t.Errorf("Expected %s quoting style", tt.expectQuote)
			}

			if !strings.Contains(result, tt.expectBool) {
				t.Errorf("Expected boolean value %s in output", tt.expectBool)
			}
		})
	}
}

func TestSQLExporterEscaping(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"Text"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Text": {Value: "It's a test", Type: models.CellTypeString, RawValue: "It's a test"},
				},
			},
		},
	}

	result, err := ToSQL(table, "test")
	if err != nil {
		t.Fatalf("ToSQL() error = %v", err)
	}

	// Single quotes should be escaped
	if !strings.Contains(result, "''") {
		t.Error("Single quotes should be escaped")
	}
}

func TestSQLExporterEmptyTable(t *testing.T) {
	table := createEmptyTable()

	result, err := ToSQL(table, "empty")
	if err != nil {
		t.Fatalf("ToSQL() error = %v", err)
	}

	// Should not contain INSERT for empty table
	if strings.Contains(result, "INSERT") {
		t.Error("Empty table should not have INSERT statement")
	}
}

func TestSQLConvenienceFunctions(t *testing.T) {
	table := createTestTable()

	// ToSQL
	result, err := ToSQL(table, "users")
	if err != nil {
		t.Errorf("ToSQL() error = %v", err)
	}
	if result == "" {
		t.Error("ToSQL() returned empty string")
	}

	// ToSQLWithCreate
	result, err = ToSQLWithCreate(table, "users")
	if err != nil {
		t.Errorf("ToSQLWithCreate() error = %v", err)
	}
	if !strings.Contains(result, "CREATE TABLE") {
		t.Error("ToSQLWithCreate() should include CREATE TABLE")
	}

	// ToSQLBuffer
	buf, err := ToSQLBuffer(table, "users")
	if err != nil {
		t.Errorf("ToSQLBuffer() error = %v", err)
	}
	if buf.Len() == 0 {
		t.Error("ToSQLBuffer() returned empty buffer")
	}
}

func TestSQLDialectString(t *testing.T) {
	tests := []struct {
		dialect  SQLDialect
		expected string
	}{
		{DialectGeneric, "generic"},
		{DialectMySQL, "mysql"},
		{DialectPostgreSQL, "postgresql"},
		{DialectSQLite, "sqlite"},
	}

	for _, tt := range tests {
		if got := tt.dialect.String(); got != tt.expected {
			t.Errorf("SQLDialect(%d).String() = %q, want %q", tt.dialect, got, tt.expected)
		}
	}
}

// ============ Edge Case Tests ============

func TestExportEmptyTable(t *testing.T) {
	table := createEmptyTable()

	// JSON
	jsonResult, _ := ToJSON(table)
	if !strings.Contains(jsonResult, `"rows":[]`) {
		t.Error("Empty table JSON should have empty rows array")
	}

	// CSV
	csvResult, _ := ToCSV(table)
	lines := strings.Split(strings.TrimSpace(csvResult), "\n")
	if len(lines) != 1 { // Only header
		t.Errorf("Empty table CSV should only have header, got %d lines", len(lines))
	}
}

func TestExportNullValues(t *testing.T) {
	table := createTestTable()

	// Test with custom null value
	opts := DefaultCSVOptions()
	opts.NullValue = "N/A"

	result, _ := NewCSVExporter(opts).ExportString(table)
	if !strings.Contains(result, "N/A") {
		t.Error("Result should contain custom null value")
	}
}

func TestSelectedColumnsNonExistent(t *testing.T) {
	table := createTestTable()
	opts := DefaultJSONOptions()
	opts.SelectedColumns = []string{"Name", "NonExistent", "Age"}

	result, err := NewJSONExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	var data map[string]interface{}
	json.Unmarshal([]byte(result), &data)

	headers := data["headers"].([]interface{})
	// Should only include existing columns
	if len(headers) != 2 {
		t.Errorf("Expected 2 headers (Name, Age), got %d", len(headers))
	}
}

func TestExportToWriter(t *testing.T) {
	table := createTestTable()
	buf := &bytes.Buffer{}

	// JSON
	err := ToJSONWriter(table, buf)
	if err != nil {
		t.Errorf("ToJSONWriter() error = %v", err)
	}

	buf.Reset()

	// CSV
	err = ToCSVWriter(table, buf)
	if err != nil {
		t.Errorf("ToCSVWriter() error = %v", err)
	}

	buf.Reset()

	// SQL
	err = ToSQLWriter(table, "test", buf)
	if err != nil {
		t.Errorf("ToSQLWriter() error = %v", err)
	}
}

// ============ Additional Coverage Tests ============

func TestCSVExporterBoolValues(t *testing.T) {
	table := createTestTable()
	result, err := ToCSV(table)
	if err != nil {
		t.Fatalf("ToCSV() error = %v", err)
	}

	if !strings.Contains(result, "true") || !strings.Contains(result, "false") {
		t.Error("CSV should contain boolean values")
	}
}

func TestSQLExporterAllTypes(t *testing.T) {
	// Test MySQL types
	opts := DefaultSQLOptions()
	opts.TableName = "test"
	opts.Dialect = DialectMySQL
	opts.CreateTable = true

	table := createTestTable()
	result, _ := NewSQLExporter(opts).ExportString(table)

	if !strings.Contains(result, "VARCHAR") {
		t.Error("MySQL should use VARCHAR")
	}

	// Test PostgreSQL types
	opts.Dialect = DialectPostgreSQL
	result, _ = NewSQLExporter(opts).ExportString(table)

	if !strings.Contains(result, "TEXT") || !strings.Contains(result, "NUMERIC") {
		t.Error("PostgreSQL should use TEXT and NUMERIC")
	}
}

func TestJSONExporterWithNullValue(t *testing.T) {
	table := createTestTable()
	opts := DefaultJSONOptions()
	opts.NullValue = "NULL_VALUE"

	result, err := NewJSONExporter(opts).ExportString(table)
	if err != nil {
		t.Fatalf("ExportString() error = %v", err)
	}

	if !strings.Contains(result, "NULL_VALUE") {
		t.Error("Result should contain custom null value")
	}
}

func TestCSVExporterWithMissingCell(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"A", "B"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"A": {Value: "value", Type: models.CellTypeString, RawValue: "value"},
					// B is missing
				},
			},
		},
	}

	result, err := ToCSV(table)
	if err != nil {
		t.Fatalf("ToCSV() error = %v", err)
	}

	// Should handle missing cell gracefully
	if result == "" {
		t.Error("Should produce output for table with missing cells")
	}
}

func TestJSONExporterWithMissingCell(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"A", "B"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"A": {Value: "value", Type: models.CellTypeString, RawValue: "value"},
					// B is missing
				},
			},
		},
	}

	result, err := ToJSON(table)
	if err != nil {
		t.Fatalf("ToJSON() error = %v", err)
	}

	if !strings.Contains(result, "null") {
		t.Error("Missing cell should be null in JSON")
	}
}

func TestSQLExporterWithMissingCell(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"A", "B"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"A": {Value: "value", Type: models.CellTypeString, RawValue: "value"},
					// B is missing
				},
			},
		},
	}

	result, err := ToSQL(table, "test")
	if err != nil {
		t.Fatalf("ToSQL() error = %v", err)
	}

	if !strings.Contains(result, "NULL") {
		t.Error("Missing cell should be NULL in SQL")
	}
}

func TestCSVExporterDefaultValue(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"A"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"A": {Value: struct{}{}, Type: models.CellTypeFormula, RawValue: "rawval"},
				},
			},
		},
	}

	result, err := ToCSV(table)
	if err != nil {
		t.Fatalf("ToCSV() error = %v", err)
	}

	// Should fall back to RawValue for unknown types
	if !strings.Contains(result, "rawval") {
		t.Error("Should use RawValue for unknown types")
	}
}

func TestSQLExporterDefaultValue(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"A"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"A": {Value: struct{}{}, Type: models.CellTypeFormula, RawValue: "rawval"},
				},
			},
		},
	}

	result, err := ToSQL(table, "test")
	if err != nil {
		t.Fatalf("ToSQL() error = %v", err)
	}

	// Should use escaped RawValue for unknown types
	if !strings.Contains(result, "'rawval'") {
		t.Error("Should use escaped RawValue for unknown types")
	}
}

func TestFilterColumnsAllColumns(t *testing.T) {
	table := createTestTable()
	headers, filter := filterColumns(table, nil)

	if len(headers) != len(table.Headers) {
		t.Errorf("Expected all headers, got %d", len(headers))
	}
	for _, h := range table.Headers {
		if !filter[h] {
			t.Errorf("Header %s should be in filter", h)
		}
	}
}

func TestGetCellValueWithNullValue(t *testing.T) {
	cell := models.Cell{Type: models.CellTypeEmpty, RawValue: ""}
	result := getCellValue(cell, "CUSTOM_NULL")
	if result != "CUSTOM_NULL" {
		t.Errorf("Expected CUSTOM_NULL, got %v", result)
	}
}

func TestGetCellValueWithoutNullValue(t *testing.T) {
	cell := models.Cell{Type: models.CellTypeEmpty, RawValue: ""}
	result := getCellValue(cell, "")
	if result != nil {
		t.Errorf("Expected nil, got %v", result)
	}
}

// ============ Error Path Tests ============

type errorWriter struct {
	err error
}

func (ew *errorWriter) Write(p []byte) (n int, err error) {
	return 0, ew.err
}

func TestJSONExporter_Export_WriteError(t *testing.T) {
	table := createTestTable()
	exporter := NewJSONExporter(nil)

	ew := &errorWriter{err: fmt.Errorf("write error")}
	err := exporter.Export(table, ew)
	if err == nil {
		t.Error("Expected error when writer fails")
	}
}

func TestCSVExporter_Export_WriteError(t *testing.T) {
	table := createTestTable()
	exporter := NewCSVExporter(nil)

	ew := &errorWriter{err: fmt.Errorf("write error")}
	err := exporter.Export(table, ew)
	if err == nil {
		t.Error("Expected error when writer fails")
	}
}

func TestSQLExporter_Export_WriteError(t *testing.T) {
	table := createTestTable()
	opts := DefaultSQLOptions()
	opts.TableName = "test"
	opts.DropTable = true
	exporter := NewSQLExporter(opts)

	ew := &errorWriter{err: fmt.Errorf("write error")}
	err := exporter.Export(table, ew)
	if err == nil {
		t.Error("Expected error when writer fails")
	}
}

func TestExport_ConvenienceWithError(t *testing.T) {
	table := createTestTable()
	ew := &errorWriter{err: fmt.Errorf("write error")}

	err := Export(table, FormatJSON, ew)
	if err == nil {
		t.Error("Expected error when writer fails")
	}
}

func TestExportString_InvalidFormat(t *testing.T) {
	table := createTestTable()
	_, err := ExportString(table, Format(99))
	if err == nil {
		t.Error("Expected error for invalid format")
	}
}

func TestCSVExporter_Export_HeaderWriteError(t *testing.T) {
	table := createTestTable()
	opts := DefaultCSVOptions()
	opts.IncludeHeaders = true
	exporter := NewCSVExporter(opts)

	// Use a writer that fails immediately
	ew := &errorWriter{err: fmt.Errorf("header write error")}
	err := exporter.Export(table, ew)
	if err == nil {
		t.Error("Expected error when header write fails")
	}
}

func TestSQLExporter_Export_CreateTableError(t *testing.T) {
	table := createTestTable()
	opts := DefaultSQLOptions()
	opts.TableName = "test"
	opts.CreateTable = true
	opts.DropTable = false
	exporter := NewSQLExporter(opts)

	ew := &errorWriter{err: fmt.Errorf("create table error")}
	err := exporter.Export(table, ew)
	if err == nil {
		t.Error("Expected error when create table write fails")
	}
}

func TestCSVExporter_FormatCell_Float64WithRawValue(t *testing.T) {
	exporter := NewCSVExporter(nil)

	// Test with RawValue present - should use RawValue
	cell := models.Cell{
		Value:    float64(3.14159),
		Type:     models.CellTypeNumber,
		RawValue: "3.14", // RawValue takes precedence
	}

	result := exporter.formatCell(cell)
	if result != "3.14" {
		t.Errorf("Expected '3.14' (from RawValue), got %q", result)
	}
}

func TestSQLExporter_InferColumnType_BoolOnly(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"Active"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Active": {Value: true, Type: models.CellTypeBool, RawValue: "true"},
				},
			},
			{
				Index: 2,
				Values: map[string]models.Cell{
					"Active": {Value: false, Type: models.CellTypeBool, RawValue: "false"},
				},
			},
		},
	}

	opts := DefaultSQLOptions()
	opts.TableName = "test"
	opts.CreateTable = true
	opts.Dialect = DialectPostgreSQL

	result, _ := NewSQLExporter(opts).ExportString(table)
	if !strings.Contains(result, "BOOLEAN") {
		t.Error("Expected BOOLEAN type for bool-only column")
	}
}

func TestSQLExporter_InferColumnType_DateOnly(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"Created"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Created": {Value: time.Now(), Type: models.CellTypeDate, RawValue: "2024-01-01"},
				},
			},
		},
	}

	opts := DefaultSQLOptions()
	opts.TableName = "test"
	opts.CreateTable = true
	opts.Dialect = DialectMySQL

	result, _ := NewSQLExporter(opts).ExportString(table)
	if !strings.Contains(result, "DATETIME") {
		t.Error("Expected DATETIME type for date-only column")
	}
}

func TestSQLExporter_InferColumnType_NumberOnly(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"Amount"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Amount": {Value: float64(100.50), Type: models.CellTypeNumber, RawValue: "100.50"},
				},
			},
		},
	}

	opts := DefaultSQLOptions()
	opts.TableName = "test"
	opts.CreateTable = true
	opts.Dialect = DialectMySQL

	result, _ := NewSQLExporter(opts).ExportString(table)
	if !strings.Contains(result, "DOUBLE") {
		t.Error("Expected DOUBLE type for number-only column in MySQL")
	}
}

func TestSQLExporter_InferColumnType_EmptyColumn(t *testing.T) {
	table := &models.Table{
		Name:    "Test",
		Headers: []string{"Empty"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Empty": {Value: nil, Type: models.CellTypeEmpty, RawValue: ""},
				},
			},
		},
	}

	opts := DefaultSQLOptions()
	opts.TableName = "test"
	opts.CreateTable = true

	result, _ := NewSQLExporter(opts).ExportString(table)
	// Empty columns should default to TEXT
	if !strings.Contains(result, "TEXT") {
		t.Error("Expected TEXT type for empty column")
	}
}

// ============ Benchmarks ============

func BenchmarkJSONExport(b *testing.B) {
	table := createTestTable()
	exporter := NewJSONExporter(nil)

	for i := 0; i < b.N; i++ {
		exporter.ExportString(table)
	}
}

func BenchmarkCSVExport(b *testing.B) {
	table := createTestTable()
	exporter := NewCSVExporter(nil)

	for i := 0; i < b.N; i++ {
		exporter.ExportString(table)
	}
}

func BenchmarkSQLExport(b *testing.B) {
	table := createTestTable()
	exporter := NewSQLExporter(nil)

	for i := 0; i < b.N; i++ {
		exporter.ExportString(table)
	}
}
