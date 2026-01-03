// Example: Context Support for Cancellation/Timeout
//
// This example demonstrates using context for timeout and cancellation
// when reading Excel files.
//
// Run: go run main.go
package main

import (
	"context"
	"errors"
	"fmt"
	"log"
	"time"

	"github.com/meddhiazoghlami/goxls"
)

func main() {
	filePath := "../../testdata/sample.xlsx"

	// Example 1: Reading with a timeout
	fmt.Println("=== Reading with Timeout ===")
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()

	workbook, err := goxls.ReadFileWithContext(ctx, filePath)
	if err != nil {
		if errors.Is(err, goxls.ErrContextCanceled) {
			fmt.Println("Operation timed out or was canceled")
		} else {
			log.Fatalf("Read failed: %v", err)
		}
	} else {
		fmt.Printf("Successfully read %d sheets\n", len(workbook.Sheets))
	}
	fmt.Println()

	// Example 2: Manual cancellation
	fmt.Println("=== Manual Cancellation ===")
	ctx2, cancel2 := context.WithCancel(context.Background())

	// Start reading in a goroutine
	done := make(chan struct{})
	go func() {
		wb, err := goxls.ReadFileWithContext(ctx2, filePath)
		if err != nil {
			if errors.Is(err, goxls.ErrContextCanceled) {
				fmt.Println("  Read was canceled")
			} else {
				fmt.Printf("  Read error: %v\n", err)
			}
		} else {
			fmt.Printf("  Read completed: %d sheets\n", len(wb.Sheets))
		}
		close(done)
	}()

	// For demonstration, we won't actually cancel since the file is small
	// In real usage, you might cancel based on user input or other conditions
	// cancel2() // Uncomment to see cancellation in action

	<-done
	cancel2() // Clean up
	fmt.Println()

	// Example 3: Combining context with options
	fmt.Println("=== Context with Options ===")
	ctx3, cancel3 := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel3()

	workbook3, err := goxls.ReadFileWithContext(ctx3, filePath,
		goxls.WithParallel(true),
		goxls.WithMinColumns(2),
	)
	if err != nil {
		log.Fatalf("Read with options failed: %v", err)
	}
	fmt.Printf("Read with options: %d sheets\n", len(workbook3.Sheets))
	fmt.Println()

	// Example 4: Error handling with sentinel errors
	fmt.Println("=== Error Handling ===")
	_, err = goxls.ReadFile("nonexistent.xlsx")
	if err != nil {
		switch {
		case errors.Is(err, goxls.ErrFileNotFound):
			fmt.Println("File not found (expected)")
		case errors.Is(err, goxls.ErrInvalidFormat):
			fmt.Println("Invalid file format")
		case errors.Is(err, goxls.ErrContextCanceled):
			fmt.Println("Operation was canceled")
		default:
			fmt.Printf("Other error: %v\n", err)
		}
	}
}
