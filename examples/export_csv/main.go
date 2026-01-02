// Example: Export to CSV
//
// This example demonstrates exporting Excel data to CSV format
// with various delimiters and options.
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

	// Method 1: Standard CSV (comma-separated)
	fmt.Println("=== Standard CSV ===")
	csv, err := goxcel.ToCSV(table)
	if err != nil {
		log.Fatalf("ToCSV failed: %v", err)
	}
	fmt.Println(csv)

	// Method 2: TSV (tab-separated)
	fmt.Println("=== TSV (Tab-Separated) ===")
	tsv, err := goxcel.ToTSV(table)
	if err != nil {
		log.Fatalf("ToTSV failed: %v", err)
	}
	fmt.Println(tsv)

	// Method 3: Custom delimiter (semicolon)
	fmt.Println("=== Semicolon-Separated ===")
	semiCSV, err := goxcel.ToCSVWithDelimiter(table, ';')
	if err != nil {
		log.Fatalf("ToCSVWithDelimiter failed: %v", err)
	}
	fmt.Println(semiCSV)

	// Method 4: Advanced options
	fmt.Println("=== Custom Options (quoted, pipe-separated) ===")
	opts := export.DefaultCSVOptions()
	opts.Delimiter = '|'
	opts.QuoteAll = true

	exporter := export.NewCSVExporter(opts)
	customCSV, err := exporter.ExportString(table)
	if err != nil {
		log.Fatalf("Custom CSV export failed: %v", err)
	}
	fmt.Println(customCSV)
}
