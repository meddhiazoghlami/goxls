// Example: Streaming large Excel files
//
// This example demonstrates how to use the streaming reader to process
// large Excel files row-by-row without loading the entire file into memory.
//
// Run: go run main.go
package main

import (
	"fmt"
	"io"
	"log"

	"github.com/meddhiazoghlami/goxls"
)

func main() {
	fmt.Println("=== Streaming Reader Example ===")
	fmt.Println()

	// Basic streaming example
	basicStreaming()

	// Streaming with options
	streamingWithOptions()

	// Using ForEach helper
	forEachExample()

	// Processing with aggregation
	aggregationExample()
}

// basicStreaming demonstrates the simplest streaming usage
func basicStreaming() {
	fmt.Println("--- Basic Streaming ---")

	// Create a streaming reader
	sr, err := goxls.NewStreamReader("../../testdata/sample.xlsx", "Sheet1")
	if err != nil {
		log.Printf("Failed to create stream reader: %v", err)
		return
	}
	defer sr.Close()

	fmt.Printf("Headers: %v\n", sr.Headers())

	// Process rows one at a time
	rowCount := 0
	for {
		row, err := sr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Printf("Error reading row: %v", err)
			break
		}

		rowCount++
		if rowCount <= 3 {
			fmt.Printf("Row %d: ", row.Index)
			for _, header := range sr.Headers() {
				if cell, ok := row.Get(header); ok {
					fmt.Printf("%s=%q ", header, cell.AsString())
				}
			}
			fmt.Println()
		}
	}

	fmt.Printf("Total rows processed: %d\n\n", rowCount)
}

// streamingWithOptions demonstrates various streaming options
func streamingWithOptions() {
	fmt.Println("--- Streaming with Options ---")

	// Skip rows, disable type detection, include empty rows
	sr, err := goxls.NewStreamReader("../../testdata/sample.xlsx", "Sheet1",
		goxls.WithStreamSkipEmptyRows(true),
		goxls.WithStreamTypeDetection(true),
		goxls.WithStreamTrimSpaces(true),
	)
	if err != nil {
		log.Printf("Failed to create stream reader: %v", err)
		return
	}
	defer sr.Close()

	row, err := sr.Next()
	if err != nil && err != io.EOF {
		log.Printf("Error: %v", err)
		return
	}

	if row != nil {
		fmt.Println("First row cell types:")
		for _, header := range sr.Headers() {
			if cell, ok := row.Get(header); ok {
				fmt.Printf("  %s: type=%v, value=%v\n", header, cell.Type, cell.Value)
			}
		}
	}
	fmt.Println()
}

// forEachExample demonstrates the ForEach helper method
func forEachExample() {
	fmt.Println("--- ForEach Helper ---")

	sr, err := goxls.NewStreamReader("../../testdata/sample.xlsx", "Sheet1")
	if err != nil {
		log.Printf("Failed to create stream reader: %v", err)
		return
	}
	defer sr.Close()

	// Process all rows with a callback
	count := 0
	err = sr.ForEach(func(row *goxls.StreamRow) error {
		count++
		return nil
	})

	if err != nil {
		log.Printf("ForEach error: %v", err)
		return
	}

	fmt.Printf("Processed %d rows using ForEach\n\n", count)
}

// aggregationExample demonstrates computing aggregates while streaming
func aggregationExample() {
	fmt.Println("--- Streaming Aggregation ---")

	sr, err := goxls.NewStreamReader("../../testdata/sample.xlsx", "Sheet1")
	if err != nil {
		log.Printf("Failed to create stream reader: %v", err)
		return
	}
	defer sr.Close()

	// Track statistics while streaming
	var (
		rowCount   int
		numericSum float64
		numericCnt int
	)

	for {
		row, err := sr.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Printf("Error: %v", err)
			break
		}

		rowCount++

		// Sum any numeric values in the row
		for _, cell := range row.Cells {
			if val, ok := cell.AsFloat(); ok {
				numericSum += val
				numericCnt++
			}
		}
	}

	fmt.Printf("Total rows: %d\n", rowCount)
	fmt.Printf("Numeric cells found: %d\n", numericCnt)
	if numericCnt > 0 {
		fmt.Printf("Sum of numeric values: %.2f\n", numericSum)
		fmt.Printf("Average: %.2f\n", numericSum/float64(numericCnt))
	}
	fmt.Println()
}
