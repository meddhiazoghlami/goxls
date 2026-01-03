package validation

import (
	"errors"
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
}
