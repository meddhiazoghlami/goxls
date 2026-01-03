package reader

import (
	"fmt"
	"sync"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// WorkbookReader is the main entry point for reading Excel files
type WorkbookReader struct {
	config         models.DetectionConfig
	analyzer       *TableAnalyzer
	headerDetector *HeaderDetector
	rowParser      *RowParser
}

// NewWorkbookReader creates a new workbook reader with default config
func NewWorkbookReader() *WorkbookReader {
	config := models.DefaultConfig()
	return &WorkbookReader{
		config:         config,
		analyzer:       NewTableAnalyzer(config),
		headerDetector: NewHeaderDetector(config),
		rowParser:      NewRowParser(config),
	}
}

// NewWorkbookReaderWithConfig creates a workbook reader with custom config
func NewWorkbookReaderWithConfig(config models.DetectionConfig) *WorkbookReader {
	return &WorkbookReader{
		config:         config,
		analyzer:       NewTableAnalyzer(config),
		headerDetector: NewHeaderDetector(config),
		rowParser:      NewRowParser(config),
	}
}

// ReadFile reads an Excel file and extracts all tables from all sheets
func (wr *WorkbookReader) ReadFile(filePath string) (*models.Workbook, error) {
	// Load the file
	excelFile, err := LoadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to load file: %w", err)
	}
	defer excelFile.Close()

	return wr.processFile(excelFile, filePath)
}

// ReadFileParallel reads an Excel file and processes sheets concurrently
// This is more efficient for workbooks with multiple sheets
func (wr *WorkbookReader) ReadFileParallel(filePath string) (*models.Workbook, error) {
	// Load the file
	excelFile, err := LoadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to load file: %w", err)
	}
	defer excelFile.Close()

	return wr.processFileParallel(excelFile, filePath)
}

// processFileParallel processes sheets concurrently
func (wr *WorkbookReader) processFileParallel(excelFile *ExcelFile, filePath string) (*models.Workbook, error) {
	sheetNames := excelFile.GetSheetNames()
	numSheets := len(sheetNames)

	if numSheets == 0 {
		return &models.Workbook{
			FilePath: filePath,
			Sheets:   []models.Sheet{},
		}, nil
	}

	// For single sheet, just process sequentially
	if numSheets == 1 {
		return wr.processFile(excelFile, filePath)
	}

	// Pre-allocate results slice to maintain sheet order
	results := make([]models.Sheet, numSheets)
	errors := make([]error, numSheets)

	var wg sync.WaitGroup
	wg.Add(numSheets)

	// Process each sheet in a goroutine
	for idx, sheetName := range sheetNames {
		go func(idx int, sheetName string) {
			defer wg.Done()

			// Each goroutine needs its own sheet processor to avoid race conditions
			// but they can share the same underlying file since reads are safe
			sheetProcessor := NewSheetProcessorWithConfig(excelFile, wr.config)

			sheet, err := wr.processSheet(sheetProcessor, sheetName, idx)
			if err != nil {
				errors[idx] = fmt.Errorf("failed to process sheet '%s': %w", sheetName, err)
				return
			}
			results[idx] = sheet
		}(idx, sheetName)
	}

	// Wait for all goroutines to complete
	wg.Wait()

	// Check for errors
	for i, err := range errors {
		if err != nil {
			return nil, fmt.Errorf("error processing sheet %d: %w", i, err)
		}
	}

	return &models.Workbook{
		FilePath: filePath,
		Sheets:   results,
	}, nil
}

// processFile processes a loaded Excel file
func (wr *WorkbookReader) processFile(excelFile *ExcelFile, filePath string) (*models.Workbook, error) {
	workbook := &models.Workbook{
		FilePath: filePath,
		Sheets:   make([]models.Sheet, 0),
	}

	// Use config-aware sheet processor for merge cell support
	sheetProcessor := NewSheetProcessorWithConfig(excelFile, wr.config)
	sheetNames := excelFile.GetSheetNames()

	for idx, sheetName := range sheetNames {
		sheet, err := wr.processSheet(sheetProcessor, sheetName, idx)
		if err != nil {
			return nil, fmt.Errorf("failed to process sheet '%s': %w", sheetName, err)
		}
		workbook.Sheets = append(workbook.Sheets, sheet)
	}

	return workbook, nil
}

// processSheet processes a single sheet and extracts tables
func (wr *WorkbookReader) processSheet(processor *SheetProcessor, sheetName string, sheetIndex int) (models.Sheet, error) {
	sheet := models.Sheet{
		Name:   sheetName,
		Index:  sheetIndex,
		Tables: make([]models.Table, 0),
	}

	// Read the sheet into a cell grid
	grid, err := processor.ReadSheet(sheetName)
	if err != nil {
		return sheet, err
	}

	if len(grid) == 0 {
		return sheet, nil
	}

	// Detect tables in the grid
	boundaries := wr.analyzer.DetectTables(grid)

	for i, boundary := range boundaries {
		table := wr.processTable(grid, boundary, sheetName, i+1)
		sheet.Tables = append(sheet.Tables, table)
	}

	return sheet, nil
}

// processTable processes a single table boundary and extracts data
func (wr *WorkbookReader) processTable(grid [][]models.Cell, boundary models.TableBoundary, sheetName string, tableNum int) models.Table {
	// Detect header row
	headerRow := wr.headerDetector.DetectHeaderRow(grid, boundary)

	// Extract headers
	headers := wr.headerDetector.ExtractHeaders(grid, headerRow, boundary)

	// Generate table name
	tableName := fmt.Sprintf("%s_Table%d", sheetName, tableNum)

	// Parse the table
	return wr.rowParser.ParseTable(grid, boundary, headers, headerRow, tableName)
}

// ReadSheet reads a single sheet by name
func (wr *WorkbookReader) ReadSheet(filePath, sheetName string) (*models.Sheet, error) {
	excelFile, err := LoadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to load file: %w", err)
	}
	defer excelFile.Close()

	// Use config-aware sheet processor for merge cell support
	sheetProcessor := NewSheetProcessorWithConfig(excelFile, wr.config)

	// Find sheet index
	sheetIndex := -1
	for i, name := range excelFile.GetSheetNames() {
		if name == sheetName {
			sheetIndex = i
			break
		}
	}

	if sheetIndex == -1 {
		return nil, fmt.Errorf("sheet '%s' not found", sheetName)
	}

	sheet, err := wr.processSheet(sheetProcessor, sheetName, sheetIndex)
	if err != nil {
		return nil, err
	}

	return &sheet, nil
}

// GetTableByName finds a table by name in a workbook
func GetTableByName(workbook *models.Workbook, tableName string) *models.Table {
	if workbook == nil {
		return nil
	}
	for _, sheet := range workbook.Sheets {
		for i := range sheet.Tables {
			if sheet.Tables[i].Name == tableName {
				return &sheet.Tables[i]
			}
		}
	}
	return nil
}

// GetAllTables returns all tables from all sheets
func GetAllTables(workbook *models.Workbook) []models.Table {
	if workbook == nil {
		return nil
	}
	var tables []models.Table
	for _, sheet := range workbook.Sheets {
		tables = append(tables, sheet.Tables...)
	}
	return tables
}
