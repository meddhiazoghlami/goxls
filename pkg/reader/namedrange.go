package reader

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/meddhiazoghlami/goxls/pkg/models"

	"github.com/xuri/excelize/v2"
)

// NamedRangeReader provides methods for working with Excel named ranges
type NamedRangeReader struct {
	config         models.DetectionConfig
	headerDetector *HeaderDetector
	rowParser      *RowParser
}

// NewNamedRangeReader creates a new named range reader
func NewNamedRangeReader() *NamedRangeReader {
	config := models.DefaultConfig()
	return &NamedRangeReader{
		config:         config,
		headerDetector: NewHeaderDetector(config),
		rowParser:      NewRowParser(config),
	}
}

// NewNamedRangeReaderWithConfig creates a new named range reader with custom config
func NewNamedRangeReaderWithConfig(config models.DetectionConfig) *NamedRangeReader {
	return &NamedRangeReader{
		config:         config,
		headerDetector: NewHeaderDetector(config),
		rowParser:      NewRowParser(config),
	}
}

// GetNamedRanges returns all named ranges from an Excel file
func (nr *NamedRangeReader) GetNamedRanges(filePath string) ([]models.NamedRange, error) {
	excelFile, err := LoadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to load file: %w", err)
	}
	defer excelFile.Close()

	return nr.getNamedRangesFromFile(excelFile), nil
}

// getNamedRangesFromFile extracts named ranges from a loaded file
func (nr *NamedRangeReader) getNamedRangesFromFile(excelFile *ExcelFile) []models.NamedRange {
	definedNames := excelFile.GetDefinedNames()
	result := make([]models.NamedRange, len(definedNames))

	for i, dn := range definedNames {
		result[i] = models.NamedRange{
			Name:     dn.Name,
			RefersTo: dn.RefersTo,
			Scope:    dn.Scope,
		}
	}

	return result
}

// ReadRange reads a named range and returns it as a table
func (nr *NamedRangeReader) ReadRange(filePath, rangeName string) (*models.Table, error) {
	excelFile, err := LoadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to load file: %w", err)
	}
	defer excelFile.Close()

	return nr.readRangeFromFile(excelFile, rangeName)
}

// readRangeFromFile reads a named range from a loaded file
func (nr *NamedRangeReader) readRangeFromFile(excelFile *ExcelFile, rangeName string) (*models.Table, error) {
	// Find the named range
	definedNames := excelFile.GetDefinedNames()
	var targetRange *DefinedNameInfo
	for i := range definedNames {
		if definedNames[i].Name == rangeName {
			targetRange = &definedNames[i]
			break
		}
	}

	if targetRange == nil {
		return nil, fmt.Errorf("named range '%s' not found", rangeName)
	}

	// Parse the reference
	sheetName, boundary, err := parseRangeReference(targetRange.RefersTo)
	if err != nil {
		return nil, fmt.Errorf("failed to parse range reference: %w", err)
	}

	// Read the sheet
	sheetProcessor := NewSheetProcessorWithConfig(excelFile, nr.config)
	grid, err := sheetProcessor.ReadSheet(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to read sheet '%s': %w", sheetName, err)
	}

	// Adjust boundary if it extends beyond grid
	if boundary.EndRow >= len(grid) {
		boundary.EndRow = len(grid) - 1
	}
	if len(grid) > 0 && boundary.EndCol >= len(grid[0]) {
		boundary.EndCol = len(grid[0]) - 1
	}

	// Validate boundary
	if boundary.StartRow > boundary.EndRow || boundary.StartCol > boundary.EndCol {
		return nil, fmt.Errorf("invalid range boundary")
	}

	// Detect header row within the boundary
	headerRow := nr.headerDetector.DetectHeaderRow(grid, boundary)

	// Extract headers
	headers := nr.headerDetector.ExtractHeaders(grid, headerRow, boundary)

	// Parse the table
	table := nr.rowParser.ParseTable(grid, boundary, headers, headerRow, rangeName)

	return &table, nil
}

// parseRangeReference parses a range reference like "Sheet1!$A$1:$B$10"
func parseRangeReference(refersTo string) (sheetName string, boundary models.TableBoundary, err error) {
	// Handle sheet name with or without quotes
	// Examples: Sheet1!$A$1:$B$10, 'Sheet Name'!$A$1:$B$10
	var cellRange string

	if strings.Contains(refersTo, "!") {
		parts := strings.SplitN(refersTo, "!", 2)
		sheetName = parts[0]
		// Remove quotes from sheet name
		sheetName = strings.Trim(sheetName, "'")
		cellRange = parts[1]
	} else {
		return "", boundary, fmt.Errorf("invalid range reference: missing sheet name")
	}

	// Parse cell range (e.g., "$A$1:$B$10" or "A1:B10")
	// Remove $ signs for simplicity
	cellRange = strings.ReplaceAll(cellRange, "$", "")

	rangeParts := strings.Split(cellRange, ":")
	if len(rangeParts) != 2 {
		return "", boundary, fmt.Errorf("invalid range format: expected 'Start:End'")
	}

	startCol, startRow, err := parseCellRef(rangeParts[0])
	if err != nil {
		return "", boundary, fmt.Errorf("invalid start cell: %w", err)
	}

	endCol, endRow, err := parseCellRef(rangeParts[1])
	if err != nil {
		return "", boundary, fmt.Errorf("invalid end cell: %w", err)
	}

	boundary = models.TableBoundary{
		StartRow: startRow,
		StartCol: startCol,
		EndRow:   endRow,
		EndCol:   endCol,
	}

	return sheetName, boundary, nil
}

// parseCellRef parses a cell reference like "A1" into column and row indices (0-based)
func parseCellRef(cellRef string) (col, row int, err error) {
	// Use excelize's built-in parser
	col, row, err = excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return 0, 0, err
	}
	// Convert to 0-based indices
	return col - 1, row - 1, nil
}

// GetNamedRangeByName finds a named range by name
func GetNamedRangeByName(ranges []models.NamedRange, name string) *models.NamedRange {
	for i := range ranges {
		if ranges[i].Name == name {
			return &ranges[i]
		}
	}
	return nil
}

// GetNamedRangesByScope returns named ranges filtered by scope
func GetNamedRangesByScope(ranges []models.NamedRange, scope string) []models.NamedRange {
	var result []models.NamedRange
	for _, r := range ranges {
		if r.Scope == scope {
			result = append(result, r)
		}
	}
	return result
}

// GetGlobalNamedRanges returns named ranges with workbook scope
func GetGlobalNamedRanges(ranges []models.NamedRange) []models.NamedRange {
	return GetNamedRangesByScope(ranges, "Workbook")
}

// ParseRangeInfo extracts sheet name and cell range from a RefersTo string
type RangeInfo struct {
	SheetName string
	StartCell string
	EndCell   string
	StartRow  int
	StartCol  int
	EndRow    int
	EndCol    int
}

// ParseNamedRangeInfo parses the RefersTo field of a named range
func ParseNamedRangeInfo(refersTo string) (*RangeInfo, error) {
	sheetName, boundary, err := parseRangeReference(refersTo)
	if err != nil {
		return nil, err
	}

	startCell, _ := excelize.CoordinatesToCellName(boundary.StartCol+1, boundary.StartRow+1)
	endCell, _ := excelize.CoordinatesToCellName(boundary.EndCol+1, boundary.EndRow+1)

	return &RangeInfo{
		SheetName: sheetName,
		StartCell: startCell,
		EndCell:   endCell,
		StartRow:  boundary.StartRow,
		StartCol:  boundary.StartCol,
		EndRow:    boundary.EndRow,
		EndCol:    boundary.EndCol,
	}, nil
}

// columnLetterToIndex converts a column letter (A, B, ..., AA, AB) to a 0-based index
// This is a fallback if excelize's parser doesn't work
var columnRegex = regexp.MustCompile(`^([A-Za-z]+)(\d+)$`)

func parseColumnRow(cellRef string) (col, row int, err error) {
	matches := columnRegex.FindStringSubmatch(cellRef)
	if matches == nil {
		return 0, 0, fmt.Errorf("invalid cell reference: %s", cellRef)
	}

	colStr := strings.ToUpper(matches[1])
	rowStr := matches[2]

	// Convert column letters to index
	col = 0
	for _, char := range colStr {
		col = col*26 + int(char-'A'+1)
	}
	col-- // Convert to 0-based

	row, err = strconv.Atoi(rowStr)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid row number: %s", rowStr)
	}
	row-- // Convert to 0-based

	return col, row, nil
}
