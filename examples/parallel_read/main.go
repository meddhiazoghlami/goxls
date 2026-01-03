// Example: Parallel Sheet Processing
//
// This example demonstrates reading Excel files with concurrent
// sheet processing for better performance on multi-sheet workbooks.
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"
	"time"

	"github.com/meddhiazoghlami/goxls"
)

func main() {
	filePath := "../../testdata/sample.xlsx"

	// Method 1: Sequential reading (default)
	fmt.Println("=== Sequential Reading ===")
	start := time.Now()
	workbook1, err := goxls.ReadFile(filePath)
	if err != nil {
		log.Fatalf("Sequential read failed: %v", err)
	}
	seqDuration := time.Since(start)
	fmt.Printf("Time: %v\n", seqDuration)
	fmt.Printf("Sheets: %d\n", len(workbook1.Sheets))
	for _, sheet := range workbook1.Sheets {
		fmt.Printf("  %s: %d tables\n", sheet.Name, len(sheet.Tables))
	}
	fmt.Println()

	// Method 2: Parallel reading
	fmt.Println("=== Parallel Reading ===")
	start = time.Now()
	workbook2, err := goxls.ReadFile(filePath, goxls.WithParallel(true))
	if err != nil {
		log.Fatalf("Parallel read failed: %v", err)
	}
	parDuration := time.Since(start)
	fmt.Printf("Time: %v\n", parDuration)
	fmt.Printf("Sheets: %d\n", len(workbook2.Sheets))
	for _, sheet := range workbook2.Sheets {
		fmt.Printf("  %s: %d tables\n", sheet.Name, len(sheet.Tables))
	}
	fmt.Println()

	// Compare results
	fmt.Println("=== Comparison ===")
	fmt.Printf("Sequential: %v\n", seqDuration)
	fmt.Printf("Parallel:   %v\n", parDuration)

	// Note: For small files, parallel may be slower due to goroutine overhead.
	// Parallel reading is most beneficial for workbooks with many sheets.
	if len(workbook1.Sheets) == 1 {
		fmt.Println("\nNote: Parallel reading provides the most benefit for")
		fmt.Println("workbooks with multiple sheets. With a single sheet,")
		fmt.Println("sequential reading may be faster due to less overhead.")
	}
}
