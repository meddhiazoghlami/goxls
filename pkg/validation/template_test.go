package validation

import (
	"errors"
	"strings"
	"testing"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// Helper to create a test workbook
func createTestWorkbook() *models.Workbook {
	return &models.Workbook{
		FilePath: "test.xlsx",
		Sheets: []models.Sheet{
			{
				Name:  "Sales",
				Index: 0,
				Tables: []models.Table{
					{
						Name:    "Sales_Table1",
						Headers: []string{"Date", "Product", "Amount", "Quantity"},
						Rows: []models.Row{
							{
								Index: 0,
								Values: map[string]models.Cell{
									"Date":     {RawValue: "2025-01-01", Type: models.CellTypeDate},
									"Product":  {RawValue: "Widget", Type: models.CellTypeString},
									"Amount":   {RawValue: "100.50", Type: models.CellTypeNumber, Value: 100.50},
									"Quantity": {RawValue: "10", Type: models.CellTypeNumber, Value: 10.0},
								},
							},
							{
								Index: 1,
								Values: map[string]models.Cell{
									"Date":     {RawValue: "2025-01-02", Type: models.CellTypeDate},
									"Product":  {RawValue: "Gadget", Type: models.CellTypeString},
									"Amount":   {RawValue: "200.00", Type: models.CellTypeNumber, Value: 200.0},
									"Quantity": {RawValue: "5", Type: models.CellTypeNumber, Value: 5.0},
								},
							},
						},
					},
				},
			},
			{
				Name:  "Inventory",
				Index: 1,
				Tables: []models.Table{
					{
						Name:    "Inventory_Table1",
						Headers: []string{"SKU", "Name", "Stock"},
						Rows: []models.Row{
							{
								Index: 0,
								Values: map[string]models.Cell{
									"SKU":   {RawValue: "SKU001", Type: models.CellTypeString},
									"Name":  {RawValue: "Widget", Type: models.CellTypeString},
									"Stock": {RawValue: "100", Type: models.CellTypeNumber, Value: 100.0},
								},
							},
						},
					},
				},
			},
		},
	}
}

func TestValidateTemplate_RequiredSheets(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: All required sheets present
	template := Template{
		RequiredSheets: []string{"Sales", "Inventory"},
	}
	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}

	// Test: Missing required sheet
	template = Template{
		RequiredSheets: []string{"Sales", "Inventory", "Missing"},
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to missing sheet")
	}
	if len(result.Errors) != 1 || result.Errors[0].Type != ErrorMissingSheet {
		t.Errorf("Expected MissingSheet error, got: %v", result.Errors)
	}
}

func TestValidateTemplate_StrictSheets(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Strict mode with unexpected sheet
	template := Template{
		RequiredSheets: []string{"Sales"},
		StrictSheets:   true,
	}
	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to unexpected sheet in strict mode")
	}

	// Test: Strict mode with all sheets accounted for
	template = Template{
		RequiredSheets: []string{"Sales"},
		OptionalSheets: []string{"Inventory"},
		StrictSheets:   true,
	}
	result = ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}
}

func TestValidateTemplate_SheetCount(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Min sheets
	template := Template{
		MinSheets: 3,
	}
	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to min sheets")
	}

	// Test: Max sheets
	template = Template{
		MaxSheets: 1,
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to max sheets")
	}

	// Test: Valid range
	template = Template{
		MinSheets: 1,
		MaxSheets: 5,
	}
	result = ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}
}

func TestValidateTemplate_RequiredColumns(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: All required columns present
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				RequiredColumns: []string{"Date", "Product", "Amount"},
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}

	// Test: Missing required column
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				RequiredColumns: []string{"Date", "Product", "Missing"},
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to missing column")
	}
}

func TestValidateTemplate_ColumnTypes(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Correct column types
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				ColumnTypes: map[string]models.CellType{
					"Amount":   models.CellTypeNumber,
					"Quantity": models.CellTypeNumber,
				},
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}

	// Test: Wrong column type
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				ColumnTypes: map[string]models.CellType{
					"Product": models.CellTypeNumber, // Product is actually String
				},
				TypeStrictness: 2, // Strict
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to wrong column type")
	}
}

func TestValidateTemplate_ColumnOrder(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Correct order
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				RequiredColumns: []string{"Date", "Product", "Amount"},
				ColumnOrder:     true,
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}

	// Test: Wrong order
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				RequiredColumns: []string{"Amount", "Date", "Product"}, // Wrong order
				ColumnOrder:     true,
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to wrong column order")
	}
}

func TestValidateTemplate_StrictColumns(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Strict columns with extra column
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				RequiredColumns: []string{"Date", "Product"},
				StrictColumns:   true,
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to unexpected columns in strict mode")
	}

	// Test: Strict columns with all columns specified
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				RequiredColumns: []string{"Date", "Product"},
				OptionalColumns: []string{"Amount", "Quantity"},
				StrictColumns:   true,
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}
}

func TestValidateTemplate_RowCount(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Min rows
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				MinRows: 10, // Sales has only 2 rows
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to min rows")
	}

	// Test: Max rows
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				MaxRows: 1, // Sales has 2 rows
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to max rows")
	}

	// Test: Valid row count
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				MinRows: 1,
				MaxRows: 10,
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}
}

func TestValidateTemplate_CustomValidation(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Custom validation passes
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				CustomValidation: func(table *models.Table) error {
					if len(table.Rows) > 0 {
						return nil
					}
					return errors.New("table must have rows")
				},
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}

	// Test: Custom validation fails
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				CustomValidation: func(table *models.Table) error {
					return errors.New("custom validation failed")
				},
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to custom validation")
	}
}

func TestValidateTemplate_AutoDetectTable(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Auto-detect first table (empty TableName)
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				RequiredColumns: []string{"Date", "Product"},
				// TableName not specified - should auto-detect
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid with auto-detect, got errors: %v", result.Errors)
	}
	if len(result.TablesValidated) == 0 {
		t.Error("Expected at least one table to be validated")
	}
}

func TestValidateTemplate_SpecificTable(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Specific table by name
	template := Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				TableName:       "Sales_Table1",
				RequiredColumns: []string{"Date"},
			},
		},
	}
	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}

	// Test: Non-existent table
	template = Template{
		RequiredSheets: []string{"Sales"},
		SheetSchemas: map[string]SheetSchema{
			"Sales": {
				TableName: "NonExistent",
			},
		},
	}
	result = ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid due to missing table")
	}
}

func TestValidateTemplate_NilWorkbook(t *testing.T) {
	template := Template{}
	result := ValidateTemplate(nil, template)
	if result.Valid {
		t.Error("Expected invalid for nil workbook")
	}
}

func TestTemplateBuilder(t *testing.T) {
	template := NewTemplate("TestTemplate").
		RequireSheets("Sales", "Inventory").
		OptionalSheets("Reports").
		StrictSheets().
		SheetCount(1, 10).
		Sheet("Sales", NewSchema().
			RequireColumns("Date", "Product").
			OptionalColumns("Notes").
			ColumnType("Amount", models.CellTypeNumber).
			ExpectOrder().
			RowCount(1, 1000).
			Build()).
		Build()

	if template.Name != "TestTemplate" {
		t.Errorf("Expected name 'TestTemplate', got '%s'", template.Name)
	}
	if len(template.RequiredSheets) != 2 { // Sales, Inventory (Sales not duplicated)
		t.Errorf("Expected 2 required sheets, got %d", len(template.RequiredSheets))
	}
	if !template.StrictSheets {
		t.Error("Expected StrictSheets to be true")
	}
	if template.MinSheets != 1 || template.MaxSheets != 10 {
		t.Error("Sheet count not set correctly")
	}

	salesSchema := template.SheetSchemas["Sales"]
	if len(salesSchema.RequiredColumns) != 2 {
		t.Errorf("Expected 2 required columns, got %d", len(salesSchema.RequiredColumns))
	}
	if !salesSchema.ColumnOrder {
		t.Error("Expected ColumnOrder to be true")
	}
}

func TestSchemaBuilder(t *testing.T) {
	schema := NewSchema().
		Table("MyTable").
		RequireColumns("A", "B", "C").
		OptionalColumns("D").
		ColumnType("A", models.CellTypeString).
		ColumnType("B", models.CellTypeNumber).
		ExpectOrder().
		StrictColumns().
		RowCount(10, 100).
		MinColumns(3).
		AllowEmpty().
		TypeStrictness(2).
		Build()

	if schema.TableName != "MyTable" {
		t.Errorf("Expected TableName 'MyTable', got '%s'", schema.TableName)
	}
	if len(schema.RequiredColumns) != 3 {
		t.Errorf("Expected 3 required columns, got %d", len(schema.RequiredColumns))
	}
	if len(schema.ColumnTypes) != 2 {
		t.Errorf("Expected 2 column types, got %d", len(schema.ColumnTypes))
	}
	if !schema.ColumnOrder {
		t.Error("Expected ColumnOrder to be true")
	}
	if !schema.StrictColumns {
		t.Error("Expected StrictColumns to be true")
	}
	if schema.MinRows != 10 || schema.MaxRows != 100 {
		t.Error("Row count not set correctly")
	}
	if !schema.AllowEmpty {
		t.Error("Expected AllowEmpty to be true")
	}
	if schema.TypeStrictness != 2 {
		t.Errorf("Expected TypeStrictness 2, got %d", schema.TypeStrictness)
	}
}

func TestQuickValidate(t *testing.T) {
	workbook := createTestWorkbook()

	// Test: Valid columns
	result := QuickValidate(workbook, "Date", "Product")
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}

	// Test: Missing column
	result = QuickValidate(workbook, "Date", "Missing")
	if result.Valid {
		t.Error("Expected invalid due to missing column")
	}

	// Test: Nil workbook
	result = QuickValidate(nil)
	if result.Valid {
		t.Error("Expected invalid for nil workbook")
	}
}

func TestValidateColumns(t *testing.T) {
	table := &models.Table{
		Headers: []string{"Date", "Product", "Amount"},
	}

	// Test: All columns present
	missing := ValidateColumns(table, "Date", "Product")
	if len(missing) != 0 {
		t.Errorf("Expected no missing columns, got: %v", missing)
	}

	// Test: Some columns missing
	missing = ValidateColumns(table, "Date", "Missing", "Also Missing")
	if len(missing) != 2 {
		t.Errorf("Expected 2 missing columns, got: %v", missing)
	}
}

func TestTemplateResult_Methods(t *testing.T) {
	result := &TemplateResult{
		Valid: false,
		Errors: []TemplateError{
			{Type: ErrorMissingSheet, Sheet: "Sheet1", Message: "Missing sheet"},
			{Type: ErrorMissingColumn, Sheet: "Sheet2", Column: "Col1", Message: "Missing column"},
			{Type: ErrorMissingColumn, Sheet: "Sheet2", Column: "Col2", Message: "Missing column"},
		},
		Warnings: []TemplateError{
			{Type: ErrorRowCount, Message: "Warning"},
		},
		SheetsValidated: []string{"Sheet1"},
		TablesValidated: []string{"Sheet1.Table1"},
	}

	if !result.HasErrors() {
		t.Error("Expected HasErrors to be true")
	}
	if !result.HasWarnings() {
		t.Error("Expected HasWarnings to be true")
	}

	byType := result.ErrorsByType()
	if len(byType[ErrorMissingColumn]) != 2 {
		t.Errorf("Expected 2 MissingColumn errors, got %d", len(byType[ErrorMissingColumn]))
	}

	bySheet := result.ErrorsBySheet()
	if len(bySheet["Sheet2"]) != 2 {
		t.Errorf("Expected 2 errors for Sheet2, got %d", len(bySheet["Sheet2"]))
	}

	summary := result.Summary()
	if summary == "" {
		t.Error("Expected non-empty summary")
	}
}

func TestTemplateErrorType_String(t *testing.T) {
	types := []TemplateErrorType{
		ErrorMissingSheet,
		ErrorUnexpectedSheet,
		ErrorSheetCount,
		ErrorMissingTable,
		ErrorMissingColumn,
		ErrorUnexpectedColumn,
		ErrorColumnOrder,
		ErrorColumnType,
		ErrorRowCount,
		ErrorColumnCount,
		ErrorCustomValidation,
	}

	for _, typ := range types {
		str := typ.String()
		if str == "" || str == "Unknown" {
			t.Errorf("Expected valid string for type %d, got '%s'", typ, str)
		}
	}

	// Test unknown type
	unknownType := TemplateErrorType(999)
	if unknownType.String() != "Unknown" {
		t.Errorf("Expected 'Unknown' for invalid type, got '%s'", unknownType.String())
	}
}

func TestTemplateError_Error(t *testing.T) {
	err := TemplateError{
		Type:    ErrorMissingColumn,
		Sheet:   "Sheet1",
		Column:  "Name",
		Message: "column not found",
	}

	errStr := err.Error()
	if errStr != "column not found" {
		t.Errorf("Expected 'column not found', got '%s'", errStr)
	}
}

func TestTemplateResult_Summary_Valid(t *testing.T) {
	result := &TemplateResult{
		Valid:           true,
		Errors:          []TemplateError{},
		Warnings:        []TemplateError{},
		SheetsValidated: []string{"Sheet1", "Sheet2"},
		TablesValidated: []string{"Sheet1.Table1", "Sheet2.Table1"},
	}

	summary := result.Summary()
	if summary == "" {
		t.Error("Expected non-empty summary")
	}
	if !strings.Contains(summary, "passed") {
		t.Error("Expected summary to contain 'passed' for valid result")
	}
	if !strings.Contains(summary, "2 sheets") {
		t.Errorf("Expected summary to mention 2 sheets, got: %s", summary)
	}
}

func TestTemplateResult_ErrorsBySheet_WorkbookLevel(t *testing.T) {
	result := &TemplateResult{
		Valid: false,
		Errors: []TemplateError{
			{Type: ErrorSheetCount, Sheet: "", Message: "Too few sheets"}, // Workbook-level error
			{Type: ErrorMissingColumn, Sheet: "Sheet1", Message: "Missing col"},
		},
	}

	bySheet := result.ErrorsBySheet()

	// Check workbook-level errors are grouped under "(workbook)"
	if len(bySheet["(workbook)"]) != 1 {
		t.Errorf("Expected 1 workbook-level error, got %d", len(bySheet["(workbook)"]))
	}
	if len(bySheet["Sheet1"]) != 1 {
		t.Errorf("Expected 1 Sheet1 error, got %d", len(bySheet["Sheet1"]))
	}
}

func TestValidateTemplate_AllTables(t *testing.T) {
	// Create workbook with multiple tables in one sheet
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Data",
				Tables: []models.Table{
					{
						Name:    "Table1",
						Headers: []string{"ID", "Name"},
						Rows: []models.Row{
							{Index: 0, Values: map[string]models.Cell{
								"ID":   {RawValue: "1", Type: models.CellTypeNumber, Value: 1.0},
								"Name": {RawValue: "Alice", Type: models.CellTypeString},
							}},
						},
					},
					{
						Name:    "Table2",
						Headers: []string{"ID", "Name"},
						Rows: []models.Row{
							{Index: 0, Values: map[string]models.Cell{
								"ID":   {RawValue: "2", Type: models.CellTypeNumber, Value: 2.0},
								"Name": {RawValue: "Bob", Type: models.CellTypeString},
							}},
						},
					},
				},
			},
		},
	}

	// Test: Validate all tables with "*"
	template := Template{
		RequiredSheets: []string{"Data"},
		SheetSchemas: map[string]SheetSchema{
			"Data": {
				TableName:       "*", // Validate all tables
				RequiredColumns: []string{"ID", "Name"},
			},
		},
	}

	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid, got errors: %v", result.Errors)
	}
	if len(result.TablesValidated) != 2 {
		t.Errorf("Expected 2 tables validated, got %d", len(result.TablesValidated))
	}
}

func TestValidateTemplate_SheetWithNoTables(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name:   "EmptySheet",
				Tables: []models.Table{}, // No tables
			},
		},
	}

	// Test: Required sheet with schema but no tables
	template := Template{
		RequiredSheets: []string{"EmptySheet"},
		SheetSchemas: map[string]SheetSchema{
			"EmptySheet": {
				RequiredColumns: []string{"ID"},
			},
		},
	}

	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid for sheet with no tables")
	}
}

func TestValidateTemplate_EmptyTableAllowed(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Sheet1",
				Tables: []models.Table{
					{
						Name:    "EmptyTable",
						Headers: []string{"ID", "Name"},
						Rows:    []models.Row{}, // Empty table
					},
				},
			},
		},
	}

	// Test: Empty table with AllowEmpty = true
	template := Template{
		SheetSchemas: map[string]SheetSchema{
			"Sheet1": {
				AllowEmpty: true,
			},
		},
	}

	result := ValidateTemplate(workbook, template)
	if !result.Valid {
		t.Errorf("Expected valid with AllowEmpty, got errors: %v", result.Errors)
	}
}

func TestValidateTemplate_EmptyTableNotAllowed(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Sheet1",
				Tables: []models.Table{
					{
						Name:    "EmptyTable",
						Headers: []string{"ID", "Name"},
						Rows:    []models.Row{}, // Empty table
					},
				},
			},
		},
	}

	// Test: Empty table with AllowEmpty = false (default)
	template := Template{
		SheetSchemas: map[string]SheetSchema{
			"Sheet1": {
				AllowEmpty: false,
			},
		},
	}

	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid for empty table without AllowEmpty")
	}
}

func TestValidateTemplate_TypeStrictnessDefault(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Sheet1",
				Tables: []models.Table{
					{
						Name:    "Table1",
						Headers: []string{"Value"},
						Rows: []models.Row{
							{Index: 0, Values: map[string]models.Cell{
								"Value": {RawValue: "100", Type: models.CellTypeNumber, Value: 100.0},
							}},
							{Index: 1, Values: map[string]models.Cell{
								"Value": {RawValue: "text", Type: models.CellTypeString, Value: "text"},
							}},
							{Index: 2, Values: map[string]models.Cell{
								"Value": {RawValue: "200", Type: models.CellTypeNumber, Value: 200.0},
							}},
						},
					},
				},
			},
		},
	}

	// Test: TypeStrictness with invalid value (uses default)
	template := Template{
		SheetSchemas: map[string]SheetSchema{
			"Sheet1": {
				ColumnTypes: map[string]models.CellType{
					"Value": models.CellTypeNumber,
				},
				TypeStrictness: 99, // Invalid value, should use default (0.5 threshold)
			},
		},
	}

	result := ValidateTemplate(workbook, template)
	// 2/3 = 66.7% match, default threshold is 50%, so should pass
	if !result.Valid {
		t.Errorf("Expected valid with default type strictness, got errors: %v", result.Errors)
	}
}

func TestValidateTemplate_MinColumns(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Sheet1",
				Tables: []models.Table{
					{
						Name:    "Table1",
						Headers: []string{"A", "B"}, // Only 2 columns
						Rows: []models.Row{
							{Index: 0, Values: map[string]models.Cell{
								"A": {RawValue: "1", Type: models.CellTypeString},
								"B": {RawValue: "2", Type: models.CellTypeString},
							}},
						},
					},
				},
			},
		},
	}

	// Test: MinColumns not met
	template := Template{
		SheetSchemas: map[string]SheetSchema{
			"Sheet1": {
				MinColumns: 5, // Requires 5 columns, table has 2
			},
		},
	}

	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid for table with too few columns")
	}
}

func TestSchemaBuilder_Custom(t *testing.T) {
	customFn := func(table *models.Table) error {
		if len(table.Rows) == 0 {
			return errors.New("table must have rows")
		}
		return nil
	}

	schema := NewSchema().
		Custom(customFn).
		Build()

	if schema.CustomValidation == nil {
		t.Error("Expected CustomValidation to be set")
	}
}

func TestValidateColumns_CaseInsensitive(t *testing.T) {
	table := &models.Table{
		Headers: []string{"ID", "Name", "EMAIL"}, // Mixed case
	}

	// Test: Case-insensitive match should work
	missing := ValidateColumns(table, "id", "name", "email")
	if len(missing) != 0 {
		t.Errorf("Expected no missing columns with case-insensitive match, got: %v", missing)
	}

	// Test: Exact match also works
	missing = ValidateColumns(table, "ID", "Name", "EMAIL")
	if len(missing) != 0 {
		t.Errorf("Expected no missing columns with exact match, got: %v", missing)
	}
}

func TestValidateTemplate_RequiredSheetWithNoSchema(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Sales",
				Tables: []models.Table{
					{Name: "Table1", Headers: []string{"A"}, Rows: []models.Row{{Index: 0}}},
				},
			},
			{
				Name:   "EmptySheet",
				Tables: []models.Table{}, // No tables
			},
		},
	}

	// Test: Required sheet without schema, but sheet has no tables
	template := Template{
		RequiredSheets: []string{"Sales", "EmptySheet"},
		// No SheetSchemas for EmptySheet
	}

	result := ValidateTemplate(workbook, template)
	// Should have a warning about EmptySheet having no tables
	if len(result.Warnings) == 0 {
		t.Error("Expected warning for required sheet with no tables")
	}
}

func TestQuickValidate_EmptyWorkbook(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{}, // Empty
	}

	result := QuickValidate(workbook, "ID")
	if result.Valid {
		t.Error("Expected invalid for empty workbook")
	}
}

func TestValidateTemplate_TypeStrictnessModerate(t *testing.T) {
	workbook := &models.Workbook{
		Sheets: []models.Sheet{
			{
				Name: "Sheet1",
				Tables: []models.Table{
					{
						Name:    "Table1",
						Headers: []string{"Value"},
						Rows: []models.Row{
							{Index: 0, Values: map[string]models.Cell{
								"Value": {RawValue: "100", Type: models.CellTypeNumber, Value: 100.0},
							}},
							{Index: 1, Values: map[string]models.Cell{
								"Value": {RawValue: "text", Type: models.CellTypeString, Value: "text"},
							}},
							{Index: 2, Values: map[string]models.Cell{
								"Value": {RawValue: "200", Type: models.CellTypeNumber, Value: 200.0},
							}},
							{Index: 3, Values: map[string]models.Cell{
								"Value": {RawValue: "300", Type: models.CellTypeNumber, Value: 300.0},
							}},
							{Index: 4, Values: map[string]models.Cell{
								"Value": {RawValue: "text2", Type: models.CellTypeString, Value: "text2"},
							}},
						},
					},
				},
			},
		},
	}

	// Test: TypeStrictness moderate (80% threshold)
	// 3/5 = 60% match, should fail with moderate
	template := Template{
		SheetSchemas: map[string]SheetSchema{
			"Sheet1": {
				ColumnTypes: map[string]models.CellType{
					"Value": models.CellTypeNumber,
				},
				TypeStrictness: 1, // Moderate (80% threshold)
			},
		},
	}

	result := ValidateTemplate(workbook, template)
	if result.Valid {
		t.Error("Expected invalid with moderate type strictness (60% < 80%)")
	}
}

func TestValidateColumns_ExactMatchWhenCaseInsensitiveFound(t *testing.T) {
	table := &models.Table{
		Headers: []string{"id", "Name"}, // lowercase "id"
	}

	// Test: Request "ID" (uppercase), case-insensitive finds "id", but exact match fails
	// The code should still find it via case-insensitive match
	missing := ValidateColumns(table, "ID")
	if len(missing) != 0 {
		t.Errorf("Expected case-insensitive match to find 'id' for 'ID', got missing: %v", missing)
	}
}
