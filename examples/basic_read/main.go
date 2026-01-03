// Example: Basic Excel file reading
//
// This example demonstrates the simplest way to read an Excel file
// using the goxls library.
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"

	"github.com/meddhiazoghlami/goxls"
)

func main() {
	// Read an Excel file with default settings
	workbook, err := goxls.ReadFile("../../testdata/sample.xlsx")
	if err != nil {
		log.Fatalf("Failed to read file: %v", err)
	}

	fmt.Printf("File: %s\n", workbook.FilePath)
	fmt.Printf("Sheets: %d\n\n", len(workbook.Sheets))

	// Iterate through all sheets
	for _, sheet := range workbook.Sheets {
		fmt.Printf("=== Sheet: %s ===\n", sheet.Name)
		fmt.Printf("Tables detected: %d\n\n", len(sheet.Tables))

		// Iterate through all tables in the sheet
		for _, table := range sheet.Tables {
			fmt.Printf("Table: %s\n", table.Name)
			fmt.Printf("  Location: Rows %d-%d, Cols %d-%d\n",
				table.StartRow, table.EndRow, table.StartCol, table.EndCol)
			fmt.Printf("  Headers: %v\n", table.Headers)
			fmt.Printf("  Row count: %d\n", table.RowCount())

			// Print first 3 rows as sample
			fmt.Println("  Sample data:")
			for i, row := range table.Rows {
				if i >= 3 {
					fmt.Printf("  ... and %d more rows\n", len(table.Rows)-3)
					break
				}

				fmt.Printf("    Row %d: ", row.Index)
				for _, header := range table.Headers {
					if cell, ok := row.Get(header); ok {
						fmt.Printf("%s=%q ", header, cell.AsString())
					}
				}
				fmt.Println()
			}
			fmt.Println()
		}
	}
}
