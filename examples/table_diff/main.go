// Example: Table Comparison (Diff)
//
// This example demonstrates comparing two tables to find differences
// (added, removed, and modified rows).
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"

	"github.com/meddhiazoghlami/goxcel"
)

func main() {
	// For this example, we'll create two versions of the same table
	// by reading the file and modifying the data

	workbook, err := goxcel.ReadFile("../../testdata/sample.xlsx")
	if err != nil {
		log.Fatalf("Failed to read file: %v", err)
	}

	if len(workbook.Sheets) == 0 || len(workbook.Sheets[0].Tables) == 0 {
		log.Fatal("No tables found in workbook")
	}

	// Use the first table as "old" version
	oldTable := &workbook.Sheets[0].Tables[0]

	// Create a "new" version by filtering some rows
	// This simulates changes to the data
	newTable := oldTable.Filter(func(row goxcel.Row) bool {
		// Keep only first half of rows (simulating removed rows)
		return row.Index < len(oldTable.Rows)/2+1
	})

	// Use first header as key column
	if len(oldTable.Headers) == 0 {
		log.Fatal("Table has no headers")
	}
	keyColumn := oldTable.Headers[0]

	fmt.Printf("Comparing tables using key column: %s\n", keyColumn)
	fmt.Printf("Old table: %d rows\n", oldTable.RowCount())
	fmt.Printf("New table: %d rows\n\n", newTable.RowCount())

	// Compare the tables
	diff := goxcel.DiffTables(oldTable, newTable, keyColumn)

	// Report results
	fmt.Println("=== Diff Results ===")
	fmt.Printf("Added rows:    %d\n", len(diff.AddedRows))
	fmt.Printf("Removed rows:  %d\n", len(diff.RemovedRows))
	fmt.Printf("Modified rows: %d\n", len(diff.ModifiedRows))
	fmt.Printf("Total changes: %d\n\n", diff.TotalChanges())

	if diff.HasChanges() {
		// Show removed rows (first 5)
		if len(diff.RemovedRows) > 0 {
			fmt.Println("=== Removed Rows ===")
			for i, row := range diff.RemovedRows {
				if i >= 5 {
					fmt.Printf("... and %d more removed rows\n", len(diff.RemovedRows)-5)
					break
				}
				if cell, ok := row.Get(keyColumn); ok {
					fmt.Printf("  Key: %s\n", cell.AsString())
				}
			}
			fmt.Println()
		}

		// Show added rows (first 5)
		if len(diff.AddedRows) > 0 {
			fmt.Println("=== Added Rows ===")
			for i, row := range diff.AddedRows {
				if i >= 5 {
					fmt.Printf("... and %d more added rows\n", len(diff.AddedRows)-5)
					break
				}
				if cell, ok := row.Get(keyColumn); ok {
					fmt.Printf("  Key: %s\n", cell.AsString())
				}
			}
			fmt.Println()
		}

		// Show modified rows with details (first 5)
		if len(diff.ModifiedRows) > 0 {
			fmt.Println("=== Modified Rows ===")
			for i, mod := range diff.ModifiedRows {
				if i >= 5 {
					fmt.Printf("... and %d more modified rows\n", len(diff.ModifiedRows)-5)
					break
				}
				fmt.Printf("  Key: %s\n", mod.KeyValue)
				for _, change := range mod.Changes {
					fmt.Printf("    %s: %q -> %q\n",
						change.Column, change.OldValue, change.NewValue)
				}
			}
		}
	} else {
		fmt.Println("No changes detected between tables.")
	}
}
