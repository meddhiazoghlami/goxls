package validation

import (
	"fmt"
	"strings"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// Template defines the expected structure of an Excel workbook.
// Use this to validate that a workbook matches your expected schema.
type Template struct {
	// Name is an optional identifier for this template
	Name string

	// RequiredSheets lists sheet names that must exist in the workbook.
	// If empty, no sheet validation is performed.
	RequiredSheets []string

	// OptionalSheets lists sheet names that may exist but aren't required.
	// These will still be validated against SheetSchemas if present.
	OptionalSheets []string

	// SheetSchemas defines the expected structure for each sheet.
	// Key is the sheet name. If a sheet is in RequiredSheets but not here,
	// only its existence is checked.
	SheetSchemas map[string]SheetSchema

	// MinSheets is the minimum number of sheets required (0 = no minimum)
	MinSheets int

	// MaxSheets is the maximum number of sheets allowed (0 = no maximum)
	MaxSheets int

	// StrictSheets when true, fails if workbook contains sheets not in
	// RequiredSheets or OptionalSheets
	StrictSheets bool
}

// SheetSchema defines the expected structure of a sheet.
type SheetSchema struct {
	// TableName specifies which table to validate in the sheet.
	// If empty, the first detected table is used.
	// Use "*" to validate all tables in the sheet.
	TableName string

	// RequiredColumns lists column names that must exist.
	RequiredColumns []string

	// OptionalColumns lists column names that may exist but aren't required.
	// These will still be validated for type if specified in ColumnTypes.
	OptionalColumns []string

	// ColumnTypes specifies expected types for columns.
	// Key is column name, value is expected CellType.
	// Only validates non-empty cells.
	ColumnTypes map[string]models.CellType

	// ColumnOrder when true, validates that RequiredColumns appear in order
	ColumnOrder bool

	// MinRows is the minimum number of data rows required (0 = no minimum)
	MinRows int

	// MaxRows is the maximum number of data rows allowed (0 = no maximum)
	MaxRows int

	// MinColumns is the minimum number of columns required (0 = no minimum)
	MinColumns int

	// StrictColumns when true, fails if table contains columns not in
	// RequiredColumns or OptionalColumns
	StrictColumns bool

	// AllowEmpty when true, allows the table to have zero rows
	AllowEmpty bool

	// TypeStrictness controls how strictly types are validated
	// 0 = lenient (default): majority of non-empty cells must match
	// 1 = moderate: 80% of non-empty cells must match
	// 2 = strict: all non-empty cells must match
	TypeStrictness int

	// CustomValidation is an optional function for custom validation logic.
	// It receives the table and should return an error if validation fails.
	CustomValidation func(table *models.Table) error
}

// TemplateError represents a single validation error.
type TemplateError struct {
	// Type categorizes the error
	Type TemplateErrorType

	// Sheet is the sheet name where the error occurred (if applicable)
	Sheet string

	// Table is the table name where the error occurred (if applicable)
	Table string

	// Column is the column name where the error occurred (if applicable)
	Column string

	// Expected is what was expected
	Expected string

	// Actual is what was found
	Actual string

	// Message is a human-readable error message
	Message string
}

// Error implements the error interface.
func (e TemplateError) Error() string {
	return e.Message
}

// TemplateErrorType categorizes template validation errors.
type TemplateErrorType int

const (
	// ErrorMissingSheet indicates a required sheet is missing
	ErrorMissingSheet TemplateErrorType = iota

	// ErrorUnexpectedSheet indicates an unexpected sheet was found (strict mode)
	ErrorUnexpectedSheet

	// ErrorSheetCount indicates sheet count is outside allowed range
	ErrorSheetCount

	// ErrorMissingTable indicates no tables were found in a sheet
	ErrorMissingTable

	// ErrorMissingColumn indicates a required column is missing
	ErrorMissingColumn

	// ErrorUnexpectedColumn indicates an unexpected column was found (strict mode)
	ErrorUnexpectedColumn

	// ErrorColumnOrder indicates columns are not in expected order
	ErrorColumnOrder

	// ErrorColumnType indicates a column has wrong data type
	ErrorColumnType

	// ErrorRowCount indicates row count is outside allowed range
	ErrorRowCount

	// ErrorColumnCount indicates column count is below minimum
	ErrorColumnCount

	// ErrorCustomValidation indicates custom validation failed
	ErrorCustomValidation
)

// String returns a string representation of the error type.
func (t TemplateErrorType) String() string {
	switch t {
	case ErrorMissingSheet:
		return "MissingSheet"
	case ErrorUnexpectedSheet:
		return "UnexpectedSheet"
	case ErrorSheetCount:
		return "SheetCount"
	case ErrorMissingTable:
		return "MissingTable"
	case ErrorMissingColumn:
		return "MissingColumn"
	case ErrorUnexpectedColumn:
		return "UnexpectedColumn"
	case ErrorColumnOrder:
		return "ColumnOrder"
	case ErrorColumnType:
		return "ColumnType"
	case ErrorRowCount:
		return "RowCount"
	case ErrorColumnCount:
		return "ColumnCount"
	case ErrorCustomValidation:
		return "CustomValidation"
	default:
		return "Unknown"
	}
}

// TemplateResult contains the results of template validation.
type TemplateResult struct {
	// Valid is true if all validations passed
	Valid bool

	// Errors contains all validation errors found
	Errors []TemplateError

	// Warnings contains non-fatal issues (e.g., optional columns missing)
	Warnings []TemplateError

	// SheetsValidated lists sheets that were validated
	SheetsValidated []string

	// TablesValidated lists tables that were validated (as "Sheet.Table")
	TablesValidated []string
}

// HasErrors returns true if there are any errors.
func (r *TemplateResult) HasErrors() bool {
	return len(r.Errors) > 0
}

// HasWarnings returns true if there are any warnings.
func (r *TemplateResult) HasWarnings() bool {
	return len(r.Warnings) > 0
}

// ErrorsByType groups errors by their type.
func (r *TemplateResult) ErrorsByType() map[TemplateErrorType][]TemplateError {
	result := make(map[TemplateErrorType][]TemplateError)
	for _, err := range r.Errors {
		result[err.Type] = append(result[err.Type], err)
	}
	return result
}

// ErrorsBySheet groups errors by sheet name.
func (r *TemplateResult) ErrorsBySheet() map[string][]TemplateError {
	result := make(map[string][]TemplateError)
	for _, err := range r.Errors {
		key := err.Sheet
		if key == "" {
			key = "(workbook)"
		}
		result[key] = append(result[key], err)
	}
	return result
}

// Summary returns a human-readable summary of the validation result.
func (r *TemplateResult) Summary() string {
	if r.Valid {
		return fmt.Sprintf("Template validation passed. Validated %d sheets, %d tables.",
			len(r.SheetsValidated), len(r.TablesValidated))
	}
	return fmt.Sprintf("Template validation failed with %d errors and %d warnings.",
		len(r.Errors), len(r.Warnings))
}

// ValidateTemplate validates a workbook against a template.
// It checks that the workbook structure matches the expected schema.
func ValidateTemplate(workbook *models.Workbook, template Template) *TemplateResult {
	result := &TemplateResult{
		Valid:           true,
		Errors:          make([]TemplateError, 0),
		Warnings:        make([]TemplateError, 0),
		SheetsValidated: make([]string, 0),
		TablesValidated: make([]string, 0),
	}

	if workbook == nil {
		result.Valid = false
		result.Errors = append(result.Errors, TemplateError{
			Type:    ErrorMissingSheet,
			Message: "workbook is nil",
		})
		return result
	}

	// Build sheet map for quick lookup
	sheetMap := make(map[string]*models.Sheet)
	for i := range workbook.Sheets {
		sheetMap[workbook.Sheets[i].Name] = &workbook.Sheets[i]
	}

	// Validate sheet count
	validateSheetCount(workbook, template, result)

	// Validate required sheets exist
	validateRequiredSheets(sheetMap, template, result)

	// Validate no unexpected sheets (strict mode)
	if template.StrictSheets {
		validateStrictSheets(workbook, template, result)
	}

	// Validate each sheet schema
	for sheetName, schema := range template.SheetSchemas {
		sheet, exists := sheetMap[sheetName]
		if !exists {
			// Already reported as missing sheet if required
			continue
		}

		result.SheetsValidated = append(result.SheetsValidated, sheetName)
		validateSheetSchema(sheet, sheetName, schema, result)
	}

	// Check if we should validate sheets without explicit schemas
	// (for required sheets not in SheetSchemas)
	for _, sheetName := range template.RequiredSheets {
		if _, hasSchema := template.SheetSchemas[sheetName]; !hasSchema {
			if sheet, exists := sheetMap[sheetName]; exists {
				// Just mark as validated since there's no schema
				if !contains(result.SheetsValidated, sheetName) {
					result.SheetsValidated = append(result.SheetsValidated, sheetName)
				}
				// Check that at least one table exists
				if len(sheet.Tables) == 0 {
					result.Warnings = append(result.Warnings, TemplateError{
						Type:    ErrorMissingTable,
						Sheet:   sheetName,
						Message: fmt.Sprintf("sheet '%s' has no detected tables", sheetName),
					})
				}
			}
		}
	}

	result.Valid = len(result.Errors) == 0
	return result
}

// validateSheetCount validates the number of sheets in the workbook.
func validateSheetCount(workbook *models.Workbook, template Template, result *TemplateResult) {
	count := len(workbook.Sheets)

	if template.MinSheets > 0 && count < template.MinSheets {
		result.Errors = append(result.Errors, TemplateError{
			Type:     ErrorSheetCount,
			Expected: fmt.Sprintf("at least %d sheets", template.MinSheets),
			Actual:   fmt.Sprintf("%d sheets", count),
			Message:  fmt.Sprintf("workbook has %d sheets, minimum required is %d", count, template.MinSheets),
		})
	}

	if template.MaxSheets > 0 && count > template.MaxSheets {
		result.Errors = append(result.Errors, TemplateError{
			Type:     ErrorSheetCount,
			Expected: fmt.Sprintf("at most %d sheets", template.MaxSheets),
			Actual:   fmt.Sprintf("%d sheets", count),
			Message:  fmt.Sprintf("workbook has %d sheets, maximum allowed is %d", count, template.MaxSheets),
		})
	}
}

// validateRequiredSheets checks that all required sheets exist.
func validateRequiredSheets(sheetMap map[string]*models.Sheet, template Template, result *TemplateResult) {
	for _, sheetName := range template.RequiredSheets {
		if _, exists := sheetMap[sheetName]; !exists {
			result.Errors = append(result.Errors, TemplateError{
				Type:     ErrorMissingSheet,
				Sheet:    sheetName,
				Expected: sheetName,
				Message:  fmt.Sprintf("required sheet '%s' not found", sheetName),
			})
		}
	}
}

// validateStrictSheets checks for unexpected sheets in strict mode.
func validateStrictSheets(workbook *models.Workbook, template Template, result *TemplateResult) {
	allowed := make(map[string]bool)
	for _, name := range template.RequiredSheets {
		allowed[name] = true
	}
	for _, name := range template.OptionalSheets {
		allowed[name] = true
	}

	for _, sheet := range workbook.Sheets {
		if !allowed[sheet.Name] {
			result.Errors = append(result.Errors, TemplateError{
				Type:    ErrorUnexpectedSheet,
				Sheet:   sheet.Name,
				Message: fmt.Sprintf("unexpected sheet '%s' found (strict mode)", sheet.Name),
			})
		}
	}
}

// validateSheetSchema validates a sheet against its schema.
func validateSheetSchema(sheet *models.Sheet, sheetName string, schema SheetSchema, result *TemplateResult) {
	// Find the table(s) to validate
	var tables []*models.Table

	if schema.TableName == "*" {
		// Validate all tables
		for i := range sheet.Tables {
			tables = append(tables, &sheet.Tables[i])
		}
	} else if schema.TableName != "" {
		// Find specific table by name
		for i := range sheet.Tables {
			if sheet.Tables[i].Name == schema.TableName {
				tables = append(tables, &sheet.Tables[i])
				break
			}
		}
		if len(tables) == 0 {
			result.Errors = append(result.Errors, TemplateError{
				Type:     ErrorMissingTable,
				Sheet:    sheetName,
				Table:    schema.TableName,
				Expected: schema.TableName,
				Message:  fmt.Sprintf("table '%s' not found in sheet '%s'", schema.TableName, sheetName),
			})
			return
		}
	} else {
		// Use first table (auto-detect)
		if len(sheet.Tables) == 0 {
			if !schema.AllowEmpty {
				result.Errors = append(result.Errors, TemplateError{
					Type:    ErrorMissingTable,
					Sheet:   sheetName,
					Message: fmt.Sprintf("no tables detected in sheet '%s'", sheetName),
				})
			}
			return
		}
		tables = append(tables, &sheet.Tables[0])
	}

	// Validate each table
	for _, table := range tables {
		tableName := fmt.Sprintf("%s.%s", sheetName, table.Name)
		result.TablesValidated = append(result.TablesValidated, tableName)
		validateTable(table, sheetName, schema, result)
	}
}

// validateTable validates a single table against a schema.
func validateTable(table *models.Table, sheetName string, schema SheetSchema, result *TemplateResult) {
	// Build header set for quick lookup
	headerSet := make(map[string]bool)
	headerIndex := make(map[string]int)
	for i, h := range table.Headers {
		headerSet[h] = true
		headerIndex[h] = i
	}

	// Validate required columns exist
	for _, col := range schema.RequiredColumns {
		if !headerSet[col] {
			result.Errors = append(result.Errors, TemplateError{
				Type:     ErrorMissingColumn,
				Sheet:    sheetName,
				Table:    table.Name,
				Column:   col,
				Expected: col,
				Message:  fmt.Sprintf("required column '%s' not found in table '%s'", col, table.Name),
			})
		}
	}

	// Validate column order (if required)
	if schema.ColumnOrder && len(schema.RequiredColumns) > 0 {
		validateColumnOrder(table, sheetName, schema, headerIndex, result)
	}

	// Validate strict columns (no unexpected columns)
	if schema.StrictColumns {
		validateStrictColumns(table, sheetName, schema, result)
	}

	// Validate column types
	if len(schema.ColumnTypes) > 0 {
		validateColumnTypes(table, sheetName, schema, result)
	}

	// Validate row count
	validateRowCount(table, sheetName, schema, result)

	// Validate column count
	if schema.MinColumns > 0 && len(table.Headers) < schema.MinColumns {
		result.Errors = append(result.Errors, TemplateError{
			Type:     ErrorColumnCount,
			Sheet:    sheetName,
			Table:    table.Name,
			Expected: fmt.Sprintf("at least %d columns", schema.MinColumns),
			Actual:   fmt.Sprintf("%d columns", len(table.Headers)),
			Message:  fmt.Sprintf("table '%s' has %d columns, minimum required is %d", table.Name, len(table.Headers), schema.MinColumns),
		})
	}

	// Run custom validation
	if schema.CustomValidation != nil {
		if err := schema.CustomValidation(table); err != nil {
			result.Errors = append(result.Errors, TemplateError{
				Type:    ErrorCustomValidation,
				Sheet:   sheetName,
				Table:   table.Name,
				Message: fmt.Sprintf("custom validation failed for table '%s': %v", table.Name, err),
			})
		}
	}
}

// validateColumnOrder checks that required columns appear in the expected order.
func validateColumnOrder(table *models.Table, sheetName string, schema SheetSchema, headerIndex map[string]int, result *TemplateResult) {
	lastIndex := -1
	for _, col := range schema.RequiredColumns {
		if idx, exists := headerIndex[col]; exists {
			if idx < lastIndex {
				result.Errors = append(result.Errors, TemplateError{
					Type:     ErrorColumnOrder,
					Sheet:    sheetName,
					Table:    table.Name,
					Column:   col,
					Expected: fmt.Sprintf("columns in order: %v", schema.RequiredColumns),
					Actual:   fmt.Sprintf("actual order: %v", table.Headers),
					Message:  fmt.Sprintf("column '%s' is out of order in table '%s'", col, table.Name),
				})
				break
			}
			lastIndex = idx
		}
	}
}

// validateStrictColumns checks for unexpected columns in strict mode.
func validateStrictColumns(table *models.Table, sheetName string, schema SheetSchema, result *TemplateResult) {
	allowed := make(map[string]bool)
	for _, col := range schema.RequiredColumns {
		allowed[col] = true
	}
	for _, col := range schema.OptionalColumns {
		allowed[col] = true
	}

	for _, header := range table.Headers {
		if !allowed[header] {
			result.Errors = append(result.Errors, TemplateError{
				Type:    ErrorUnexpectedColumn,
				Sheet:   sheetName,
				Table:   table.Name,
				Column:  header,
				Message: fmt.Sprintf("unexpected column '%s' in table '%s' (strict mode)", header, table.Name),
			})
		}
	}
}

// validateColumnTypes checks that column values match expected types.
func validateColumnTypes(table *models.Table, sheetName string, schema SheetSchema, result *TemplateResult) {
	for colName, expectedType := range schema.ColumnTypes {
		// Count type matches
		total := 0
		matches := 0

		for _, row := range table.Rows {
			if cell, ok := row.Get(colName); ok {
				if !cell.IsEmpty() {
					total++
					if cell.Type == expectedType {
						matches++
					}
				}
			}
		}

		if total == 0 {
			// No non-empty values to check
			continue
		}

		// Calculate match percentage
		matchRate := float64(matches) / float64(total)

		// Determine threshold based on strictness
		var threshold float64
		switch schema.TypeStrictness {
		case 0: // lenient
			threshold = 0.5
		case 1: // moderate
			threshold = 0.8
		case 2: // strict
			threshold = 1.0
		default:
			threshold = 0.5
		}

		if matchRate < threshold {
			result.Errors = append(result.Errors, TemplateError{
				Type:     ErrorColumnType,
				Sheet:    sheetName,
				Table:    table.Name,
				Column:   colName,
				Expected: fmt.Sprintf("type %v (%.0f%% threshold)", expectedType, threshold*100),
				Actual:   fmt.Sprintf("%.1f%% match (%d/%d)", matchRate*100, matches, total),
				Message: fmt.Sprintf("column '%s' type mismatch: expected %v, got %.1f%% match",
					colName, expectedType, matchRate*100),
			})
		}
	}
}

// validateRowCount checks that the table has the expected number of rows.
func validateRowCount(table *models.Table, sheetName string, schema SheetSchema, result *TemplateResult) {
	count := len(table.Rows)

	if !schema.AllowEmpty && count == 0 {
		result.Errors = append(result.Errors, TemplateError{
			Type:    ErrorRowCount,
			Sheet:   sheetName,
			Table:   table.Name,
			Message: fmt.Sprintf("table '%s' is empty (no data rows)", table.Name),
		})
		return
	}

	if schema.MinRows > 0 && count < schema.MinRows {
		result.Errors = append(result.Errors, TemplateError{
			Type:     ErrorRowCount,
			Sheet:    sheetName,
			Table:    table.Name,
			Expected: fmt.Sprintf("at least %d rows", schema.MinRows),
			Actual:   fmt.Sprintf("%d rows", count),
			Message:  fmt.Sprintf("table '%s' has %d rows, minimum required is %d", table.Name, count, schema.MinRows),
		})
	}

	if schema.MaxRows > 0 && count > schema.MaxRows {
		result.Errors = append(result.Errors, TemplateError{
			Type:     ErrorRowCount,
			Sheet:    sheetName,
			Table:    table.Name,
			Expected: fmt.Sprintf("at most %d rows", schema.MaxRows),
			Actual:   fmt.Sprintf("%d rows", count),
			Message:  fmt.Sprintf("table '%s' has %d rows, maximum allowed is %d", table.Name, count, schema.MaxRows),
		})
	}
}

// contains checks if a string slice contains a value.
func contains(slice []string, value string) bool {
	for _, v := range slice {
		if v == value {
			return true
		}
	}
	return false
}

// --- Builder API for easier template creation ---

// TemplateBuilder provides a fluent API for building templates.
type TemplateBuilder struct {
	template Template
}

// NewTemplate creates a new template builder.
func NewTemplate(name string) *TemplateBuilder {
	return &TemplateBuilder{
		template: Template{
			Name:         name,
			SheetSchemas: make(map[string]SheetSchema),
		},
	}
}

// RequireSheets adds required sheets to the template.
func (b *TemplateBuilder) RequireSheets(names ...string) *TemplateBuilder {
	b.template.RequiredSheets = append(b.template.RequiredSheets, names...)
	return b
}

// OptionalSheets adds optional sheets to the template.
func (b *TemplateBuilder) OptionalSheets(names ...string) *TemplateBuilder {
	b.template.OptionalSheets = append(b.template.OptionalSheets, names...)
	return b
}

// StrictSheets enables strict sheet validation.
func (b *TemplateBuilder) StrictSheets() *TemplateBuilder {
	b.template.StrictSheets = true
	return b
}

// SheetCount sets the allowed sheet count range.
func (b *TemplateBuilder) SheetCount(min, max int) *TemplateBuilder {
	b.template.MinSheets = min
	b.template.MaxSheets = max
	return b
}

// Sheet adds or updates a sheet schema.
func (b *TemplateBuilder) Sheet(name string, schema SheetSchema) *TemplateBuilder {
	b.template.SheetSchemas[name] = schema
	// Auto-add to required sheets if not already there
	if !contains(b.template.RequiredSheets, name) && !contains(b.template.OptionalSheets, name) {
		b.template.RequiredSheets = append(b.template.RequiredSheets, name)
	}
	return b
}

// Build returns the constructed template.
func (b *TemplateBuilder) Build() Template {
	return b.template
}

// --- SheetSchema Builder ---

// SchemaBuilder provides a fluent API for building sheet schemas.
type SchemaBuilder struct {
	schema SheetSchema
}

// NewSchema creates a new sheet schema builder.
func NewSchema() *SchemaBuilder {
	return &SchemaBuilder{
		schema: SheetSchema{
			ColumnTypes: make(map[string]models.CellType),
		},
	}
}

// Table specifies which table to validate.
// Use "" for auto-detect (first table), "*" for all tables.
func (b *SchemaBuilder) Table(name string) *SchemaBuilder {
	b.schema.TableName = name
	return b
}

// RequireColumns adds required columns.
func (b *SchemaBuilder) RequireColumns(names ...string) *SchemaBuilder {
	b.schema.RequiredColumns = append(b.schema.RequiredColumns, names...)
	return b
}

// OptionalColumns adds optional columns.
func (b *SchemaBuilder) OptionalColumns(names ...string) *SchemaBuilder {
	b.schema.OptionalColumns = append(b.schema.OptionalColumns, names...)
	return b
}

// ColumnType specifies the expected type for a column.
func (b *SchemaBuilder) ColumnType(name string, cellType models.CellType) *SchemaBuilder {
	b.schema.ColumnTypes[name] = cellType
	return b
}

// ExpectOrder enables column order validation.
func (b *SchemaBuilder) ExpectOrder() *SchemaBuilder {
	b.schema.ColumnOrder = true
	return b
}

// StrictColumns enables strict column validation.
func (b *SchemaBuilder) StrictColumns() *SchemaBuilder {
	b.schema.StrictColumns = true
	return b
}

// RowCount sets the allowed row count range.
func (b *SchemaBuilder) RowCount(min, max int) *SchemaBuilder {
	b.schema.MinRows = min
	b.schema.MaxRows = max
	return b
}

// MinColumns sets the minimum column count.
func (b *SchemaBuilder) MinColumns(min int) *SchemaBuilder {
	b.schema.MinColumns = min
	return b
}

// AllowEmpty allows the table to have zero rows.
func (b *SchemaBuilder) AllowEmpty() *SchemaBuilder {
	b.schema.AllowEmpty = true
	return b
}

// TypeStrictness sets the type validation strictness (0=lenient, 1=moderate, 2=strict).
func (b *SchemaBuilder) TypeStrictness(level int) *SchemaBuilder {
	b.schema.TypeStrictness = level
	return b
}

// Custom adds a custom validation function.
func (b *SchemaBuilder) Custom(fn func(*models.Table) error) *SchemaBuilder {
	b.schema.CustomValidation = fn
	return b
}

// Build returns the constructed schema.
func (b *SchemaBuilder) Build() SheetSchema {
	return b.schema
}

// --- Quick validation helpers ---

// QuickValidate performs a simple validation with just required columns.
// This is a convenience function for basic validation needs.
func QuickValidate(workbook *models.Workbook, requiredColumns ...string) *TemplateResult {
	if workbook == nil || len(workbook.Sheets) == 0 {
		return &TemplateResult{
			Valid:  false,
			Errors: []TemplateError{{Type: ErrorMissingSheet, Message: "workbook is empty or nil"}},
		}
	}

	// Use first sheet, first table
	schema := NewSchema().RequireColumns(requiredColumns...).Build()
	template := NewTemplate("QuickValidation").
		Sheet(workbook.Sheets[0].Name, schema).
		Build()

	return ValidateTemplate(workbook, template)
}

// ValidateColumns validates that a table has the required columns.
func ValidateColumns(table *models.Table, requiredColumns ...string) []string {
	missing := make([]string, 0)
	headerSet := make(map[string]bool)
	for _, h := range table.Headers {
		headerSet[strings.ToLower(h)] = true
	}

	for _, col := range requiredColumns {
		if !headerSet[strings.ToLower(col)] {
			// Try exact match
			found := false
			for _, h := range table.Headers {
				if h == col {
					found = true
					break
				}
			}
			if !found {
				missing = append(missing, col)
			}
		}
	}

	return missing
}
