// Example: Template Validation
//
// This example demonstrates validating Excel workbook structure against
// a predefined template using the template validation feature.
// This validates sheet names, column names, types, and counts.
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"

	"github.com/meddhiazoghlami/goxls"
)

func main() {
	// Read Excel file
	workbook, err := goxls.ReadFile("../../testdata/sample.xlsx")
	if err != nil {
		log.Fatalf("Failed to read file: %v", err)
	}

	fmt.Printf("Loaded workbook with %d sheet(s)\n", len(workbook.Sheets))
	for _, sheet := range workbook.Sheets {
		fmt.Printf("  - %s (%d tables)\n", sheet.Name, len(sheet.Tables))
		for _, table := range sheet.Tables {
			fmt.Printf("      Table '%s': %v\n", table.Name, table.Headers)
		}
	}
	fmt.Println()

	// === Example 1: Quick Validation ===
	// Simple validation to check if required columns exist
	fmt.Println("=== Example 1: Quick Validation ===")
	quickResult := goxls.QuickValidate(workbook, "Name", "Value")
	if quickResult.Valid {
		fmt.Println("Quick validation passed!")
	} else {
		fmt.Printf("Quick validation failed: %d errors\n", len(quickResult.Errors))
		for _, err := range quickResult.Errors {
			fmt.Printf("  - %s\n", err.Message)
		}
	}
	fmt.Println()

	// === Example 2: Template with Required Sheets ===
	fmt.Println("=== Example 2: Template with Required Sheets ===")
	template2 := goxls.NewTemplate("SheetCheck").
		RequireSheets("Sheet1").
		Build()

	result2 := goxls.ValidateTemplate(workbook, template2)
	fmt.Printf("Result: %s\n", result2.Summary())
	fmt.Println()

	// === Example 3: Full Template with Schema ===
	fmt.Println("=== Example 3: Full Template with Schema ===")

	// Define expected schema for a sheet
	// Auto-detects the first table when TableName is empty
	schema := goxls.NewSchema().
		RequireColumns("Name").        // These columns must exist
		OptionalColumns("Value").      // These may or may not exist
		RowCount(1, 100).              // Expect 1-100 data rows
		Build()

	template3 := goxls.NewTemplate("FullTemplate").
		RequireSheets("Sheet1").
		Sheet("Sheet1", schema).
		Build()

	result3 := goxls.ValidateTemplate(workbook, template3)
	fmt.Printf("Result: %s\n", result3.Summary())
	if !result3.Valid {
		fmt.Println("Errors:")
		for _, err := range result3.Errors {
			fmt.Printf("  - [%s] %s\n", err.Type.String(), err.Message)
		}
	}
	fmt.Println()

	// === Example 4: Strict Mode Validation ===
	fmt.Println("=== Example 4: Strict Mode (no extra sheets/columns) ===")

	strictSchema := goxls.NewSchema().
		RequireColumns("Name", "Value").
		StrictColumns(). // Fail if unexpected columns exist
		Build()

	template4 := goxls.NewTemplate("StrictTemplate").
		RequireSheets("Sheet1").
		StrictSheets(). // Fail if unexpected sheets exist
		Sheet("Sheet1", strictSchema).
		Build()

	result4 := goxls.ValidateTemplate(workbook, template4)
	fmt.Printf("Result: %s\n", result4.Summary())
	if !result4.Valid {
		fmt.Println("Errors (expected in strict mode):")
		for _, err := range result4.Errors {
			fmt.Printf("  - [%s] %s\n", err.Type.String(), err.Message)
		}
	}
	fmt.Println()

	// === Example 5: Column Type Validation ===
	fmt.Println("=== Example 5: Column Type Validation ===")

	typeSchema := goxls.NewSchema().
		RequireColumns("Name", "Value").
		ColumnType("Name", goxls.CellTypeString).
		ColumnType("Value", goxls.CellTypeNumber).
		TypeStrictness(goxls.TypeStrictnessLenient). // 50% threshold
		Build()

	template5 := goxls.NewTemplate("TypeTemplate").
		Sheet("Sheet1", typeSchema).
		Build()

	result5 := goxls.ValidateTemplate(workbook, template5)
	fmt.Printf("Result: %s\n", result5.Summary())
	if !result5.Valid {
		for _, err := range result5.Errors {
			if err.Type == goxls.ErrorColumnType {
				fmt.Printf("  - Type mismatch: %s (expected: %s, actual: %s)\n",
					err.Column, err.Expected, err.Actual)
			}
		}
	}
	fmt.Println()

	// === Example 6: Column Order Validation ===
	fmt.Println("=== Example 6: Column Order Validation ===")

	orderSchema := goxls.NewSchema().
		RequireColumns("Name", "Value"). // Must appear in this order
		ExpectOrder().
		Build()

	template6 := goxls.NewTemplate("OrderTemplate").
		Sheet("Sheet1", orderSchema).
		Build()

	result6 := goxls.ValidateTemplate(workbook, template6)
	fmt.Printf("Result: %s\n", result6.Summary())
	fmt.Println()

	// === Example 7: ValidateColumns Helper ===
	fmt.Println("=== Example 7: ValidateColumns Helper ===")
	if len(workbook.Sheets) > 0 && len(workbook.Sheets[0].Tables) > 0 {
		table := &workbook.Sheets[0].Tables[0]
		missing := goxls.ValidateColumns(table, "Name", "Value", "NonExistentColumn")
		if len(missing) > 0 {
			fmt.Printf("Missing columns: %v\n", missing)
		} else {
			fmt.Println("All columns present!")
		}
	}
	fmt.Println()

	// === Example 8: Using Template Struct Directly ===
	fmt.Println("=== Example 8: Direct Template Struct ===")

	directTemplate := goxls.Template{
		Name:           "DirectTemplate",
		RequiredSheets: []string{"Sheet1"},
		SheetSchemas: map[string]goxls.SheetSchema{
			"Sheet1": {
				RequiredColumns: []string{"Name"},
				MinRows:         1,
				AllowEmpty:      false,
			},
		},
	}

	result8 := goxls.ValidateTemplate(workbook, directTemplate)
	fmt.Printf("Result: %s\n", result8.Summary())
	fmt.Printf("Sheets validated: %v\n", result8.SheetsValidated)
	fmt.Printf("Tables validated: %v\n", result8.TablesValidated)
}
