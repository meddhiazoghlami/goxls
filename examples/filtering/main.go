// Example: Row Filtering and Deduplication
//
// This example demonstrates filtering rows, selecting columns,
// and removing duplicates from Excel tables.
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"

	"github.com/meddhiazoghlami/goxcel"
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

	fmt.Printf("Original table: %s\n", table.Name)
	fmt.Printf("  Rows: %d\n", table.RowCount())
	fmt.Printf("  Headers: %v\n\n", table.Headers)

	// Example 1: Filter rows (keep rows where first column is not empty)
	fmt.Println("=== Filtering: Non-empty first column ===")
	if len(table.Headers) > 0 {
		firstCol := table.Headers[0]
		filtered := table.Filter(func(row goxcel.Row) bool {
			if cell, ok := row.Get(firstCol); ok {
				return !cell.IsEmpty()
			}
			return false
		})
		fmt.Printf("Filtered rows: %d (was %d)\n\n", filtered.RowCount(), table.RowCount())
	}

	// Example 2: Select specific columns
	fmt.Println("=== Column Selection ===")
	if len(table.Headers) >= 2 {
		selected := table.Select(table.Headers[0], table.Headers[1])
		fmt.Printf("Selected columns: %v\n", selected.Headers)
		fmt.Printf("Row count: %d\n\n", selected.RowCount())
	}

	// Example 3: Rename columns
	fmt.Println("=== Column Renaming ===")
	if len(table.Headers) > 0 {
		renamed := table.Rename(map[string]string{
			table.Headers[0]: "FirstColumn",
		})
		fmt.Printf("Renamed headers: %v\n\n", renamed.Headers)
	}

	// Example 4: Reorder columns
	fmt.Println("=== Column Reordering ===")
	if len(table.Headers) >= 2 {
		// Reverse first two columns
		reordered := table.Reorder(table.Headers[1], table.Headers[0])
		fmt.Printf("Reordered headers: %v\n\n", reordered.Headers)
	}

	// Example 5: Find duplicates
	fmt.Println("=== Duplicate Detection ===")
	if len(table.Headers) > 0 {
		keyCol := table.Headers[0]
		duplicates := table.FindDuplicates(keyCol)
		fmt.Printf("Duplicate rows (key=%s): %d\n", keyCol, len(duplicates))

		groups := table.FindDuplicateGroups(keyCol)
		if len(groups) > 0 {
			fmt.Println("Duplicate groups:")
			for i, group := range groups {
				if i >= 5 {
					fmt.Printf("  ... and %d more groups\n", len(groups)-5)
					break
				}
				fmt.Printf("  '%s': %d occurrences\n", group.KeyValue, group.Count)
			}
		}
		fmt.Println()
	}

	// Example 6: Deduplicate
	fmt.Println("=== Deduplication ===")
	if len(table.Headers) > 0 {
		keyCol := table.Headers[0]
		unique := table.Deduplicate(keyCol)
		fmt.Printf("After deduplication: %d rows (was %d)\n", unique.RowCount(), table.RowCount())
	}

	// Example 7: Chained operations
	fmt.Println("\n=== Chained Operations ===")
	if len(table.Headers) >= 2 {
		result := table.
			Filter(func(row goxcel.Row) bool {
				// Keep all rows for this demo
				return true
			}).
			Select(table.Headers[0], table.Headers[1]).
			Deduplicate(table.Headers[0])

		fmt.Printf("After chain: %d rows, %d columns\n",
			result.RowCount(), len(result.Headers))
	}
}
