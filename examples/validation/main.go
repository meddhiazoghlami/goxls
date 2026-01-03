// Example: Data Validation
//
// This example demonstrates validating Excel table data against
// custom rules using the validation package.
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"

	"github.com/meddhiazoghlami/goxls"
	"github.com/meddhiazoghlami/goxls/pkg/validation"
)

func main() {
	// Read Excel file
	workbook, err := goxls.ReadFile("../../testdata/sample.xlsx")
	if err != nil {
		log.Fatalf("Failed to read file: %v", err)
	}

	if len(workbook.Sheets) == 0 || len(workbook.Sheets[0].Tables) == 0 {
		log.Fatal("No tables found in workbook")
	}

	table := &workbook.Sheets[0].Tables[0]

	fmt.Printf("Validating table: %s (%d rows)\n", table.Name, table.RowCount())
	fmt.Printf("Headers: %v\n\n", table.Headers)

	// Define validation rules using the fluent builder API
	// Adjust these based on your actual data columns
	rules := []validation.ValidationRule{}

	// Add rules based on available headers
	for _, header := range table.Headers {
		// Make all columns required as a basic check
		rules = append(rules,
			validation.ForColumn(header).Required().Build(),
		)
	}

	// Validate the table
	result := validation.ValidateTable(table, rules)

	// Report results
	if result.Valid {
		fmt.Println("Validation PASSED - all rows are valid!")
	} else {
		fmt.Printf("Validation FAILED - %d errors found\n\n", len(result.Errors))

		// Show first 10 errors
		for i, verr := range result.Errors {
			if i >= 10 {
				fmt.Printf("... and %d more errors\n", len(result.Errors)-10)
				break
			}
			fmt.Printf("  Row %d, Column '%s': %s\n", verr.Row, verr.Column, verr.Message)
		}

		// Group errors by column
		fmt.Println("\n=== Errors by Column ===")
		byColumn := result.ErrorsByColumn()
		for col, errors := range byColumn {
			fmt.Printf("  %s: %d errors\n", col, len(errors))
		}
	}
}
