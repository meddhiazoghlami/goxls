// Package reader provides functionality for reading Excel files (.xlsx) with
// automatic table detection and intelligent data extraction.
//
// The package automatically detects table boundaries, identifies headers,
// and extracts structured data without requiring predefined schemas.
//
// # Basic Usage
//
// The simplest way to read an Excel file:
//
//	wr := reader.NewWorkbookReader()
//	workbook, err := wr.ReadFile("data.xlsx")
//	if err != nil {
//	    log.Fatal(err)
//	}
//
//	for _, sheet := range workbook.Sheets {
//	    for _, table := range sheet.Tables {
//	        fmt.Printf("Table: %s with %d rows\n", table.Name, len(table.Rows))
//	    }
//	}
//
// # Custom Configuration
//
// Configure table detection parameters:
//
//	config := models.DetectionConfig{
//	    MinColumns:        3,     // Minimum columns for table detection
//	    MinRows:           5,     // Minimum rows for table detection
//	    MaxEmptyRows:      2,     // Max empty rows before table boundary
//	    HeaderDensity:     0.6,   // Min density for header row
//	    ExpandMergedCells: true,  // Copy merged cell values
//	}
//	wr := reader.NewWorkbookReaderWithConfig(config)
//
// # Parallel Processing
//
// For workbooks with multiple sheets, use parallel processing:
//
//	workbook, err := wr.ReadFileParallel("multi_sheet.xlsx")
//
// # Named Ranges
//
// Read Excel named ranges:
//
//	nr := reader.NewNamedRangeReader()
//	ranges, _ := nr.GetNamedRanges("data.xlsx")
//	table, _ := nr.ReadRange("data.xlsx", "SalesData")
//
// # Components
//
// The reader package consists of several components:
//
//   - WorkbookReader: Main entry point for reading Excel files
//   - SheetProcessor: Reads individual sheets into cell grids
//   - TableAnalyzer: Detects table boundaries using density analysis
//   - HeaderDetector: Identifies header rows with scoring algorithms
//   - RowParser: Converts grid rows to structured Row objects
//   - MergeProcessor: Handles merged cell regions
//   - NamedRangeReader: Reads Excel named ranges
package reader
