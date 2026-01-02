// Example: Export to JSON
//
// This example demonstrates exporting Excel data to JSON format
// with various options.
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"

	"github.com/meddhiazoghlami/goxcel"
	"github.com/meddhiazoghlami/goxcel/pkg/export"
)

func main() {
	// Read Excel file
	workbook, err := goxcel.ReadFile("../../testdata/sample.xlsx")
	if err != nil {
		log.Fatalf("Failed to read file: %v", err)
	}

	if len(workbook.Sheets) == 0 || len(workbook.Sheets[0].Tables) == 0 {
		log.Fatal("No tables found in workbook")
	}

	table := &workbook.Sheets[0].Tables[0]

	// Method 1: Simple JSON export (compact)
	fmt.Println("=== Compact JSON ===")
	json, err := goxcel.ToJSON(table)
	if err != nil {
		log.Fatalf("ToJSON failed: %v", err)
	}
	fmt.Println(json)
	fmt.Println()

	// Method 2: Pretty-printed JSON
	fmt.Println("=== Pretty JSON ===")
	jsonPretty, err := goxcel.ToJSONPretty(table)
	if err != nil {
		log.Fatalf("ToJSONPretty failed: %v", err)
	}
	fmt.Println(jsonPretty)
	fmt.Println()

	// Method 3: JSON with custom options (select specific columns)
	fmt.Println("=== Custom JSON (selected columns) ===")
	opts := export.DefaultJSONOptions()
	opts.Pretty = true
	// Select only first two columns if available
	if len(table.Headers) >= 2 {
		opts.SelectedColumns = table.Headers[:2]
	}

	exporter := export.NewJSONExporter(opts)
	customJSON, err := exporter.ExportString(table)
	if err != nil {
		log.Fatalf("Custom JSON export failed: %v", err)
	}
	fmt.Println(customJSON)
}
