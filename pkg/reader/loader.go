package reader

import (
	"errors"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

var (
	ErrFileNotFound    = errors.New("file not found")
	ErrInvalidFormat   = errors.New("invalid file format: only .xlsx files are supported")
	ErrFileEmpty       = errors.New("file is empty")
	ErrCannotOpenFile  = errors.New("cannot open file")
)

// ExcelFile wraps an excelize file with additional functionality
type ExcelFile struct {
	file     *excelize.File
	filePath string
}

// LoadFile opens and validates an Excel file
func LoadFile(path string) (*ExcelFile, error) {
	// Check if file exists
	info, err := os.Stat(path)
	if os.IsNotExist(err) {
		return nil, ErrFileNotFound
	}
	if err != nil {
		return nil, err
	}

	// Check if file is empty
	if info.Size() == 0 {
		return nil, ErrFileEmpty
	}

	// Check file extension
	ext := filepath.Ext(path)
	if ext != ".xlsx" {
		return nil, ErrInvalidFormat
	}

	// Open the file
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, errors.Join(ErrCannotOpenFile, err)
	}

	return &ExcelFile{
		file:     f,
		filePath: path,
	}, nil
}

// Close closes the Excel file
func (ef *ExcelFile) Close() error {
	if ef.file != nil {
		return ef.file.Close()
	}
	return nil
}

// GetSheetNames returns all sheet names in the workbook
func (ef *ExcelFile) GetSheetNames() []string {
	return ef.file.GetSheetList()
}

// GetSheetCount returns the number of sheets
func (ef *ExcelFile) GetSheetCount() int {
	return len(ef.file.GetSheetList())
}

// GetRows returns all rows for a given sheet
func (ef *ExcelFile) GetRows(sheetName string) ([][]string, error) {
	return ef.file.GetRows(sheetName)
}

// GetCellValue returns the value of a specific cell
func (ef *ExcelFile) GetCellValue(sheetName string, cell string) (string, error) {
	return ef.file.GetCellValue(sheetName, cell)
}

// GetCellType returns the type of a specific cell
func (ef *ExcelFile) GetCellType(sheetName string, cell string) (excelize.CellType, error) {
	return ef.file.GetCellType(sheetName, cell)
}

// FilePath returns the path of the loaded file
func (ef *ExcelFile) FilePath() string {
	return ef.filePath
}

// Raw returns the underlying excelize.File for advanced operations
func (ef *ExcelFile) Raw() *excelize.File {
	return ef.file
}

// MergeCellInfo represents a merged cell region
type MergeCellInfo struct {
	StartCell string // Top-left cell reference (e.g., "A1")
	EndCell   string // Bottom-right cell reference (e.g., "C3")
	Value     string // The merged cell's value
}

// GetMergeCells returns all merged cell regions in a sheet
func (ef *ExcelFile) GetMergeCells(sheetName string) ([]MergeCellInfo, error) {
	mergeCells, err := ef.file.GetMergeCells(sheetName)
	if err != nil {
		return nil, err
	}

	result := make([]MergeCellInfo, len(mergeCells))
	for i, mc := range mergeCells {
		result[i] = MergeCellInfo{
			StartCell: mc.GetStartAxis(),
			EndCell:   mc.GetEndAxis(),
			Value:     mc.GetCellValue(),
		}
	}
	return result, nil
}

// GetCellFormula returns the formula for a specific cell, or empty string if no formula
func (ef *ExcelFile) GetCellFormula(sheetName string, cell string) (string, error) {
	return ef.file.GetCellFormula(sheetName, cell)
}

// CommentInfo represents a cell comment
type CommentInfo struct {
	Cell   string // Cell reference (e.g., "A1")
	Author string // Comment author
	Text   string // Comment text
}

// GetComments returns all comments for a sheet
func (ef *ExcelFile) GetComments(sheetName string) ([]CommentInfo, error) {
	comments, err := ef.file.GetComments(sheetName)
	if err != nil {
		return nil, err
	}

	result := make([]CommentInfo, len(comments))
	for i, c := range comments {
		result[i] = CommentInfo{
			Cell:   c.Cell,
			Author: c.Author,
			Text:   c.Text,
		}
	}
	return result, nil
}

// GetCellHyperLink returns the hyperlink URL for a specific cell, or empty string if none
func (ef *ExcelFile) GetCellHyperLink(sheetName string, cell string) (string, error) {
	hasLink, link, err := ef.file.GetCellHyperLink(sheetName, cell)
	if err != nil {
		return "", err
	}
	if !hasLink {
		return "", nil
	}
	return link, nil
}

// DefinedNameInfo represents an Excel named range
type DefinedNameInfo struct {
	Name     string // The name of the range
	RefersTo string // The cell reference (e.g., "Sheet1!$A$1:$B$10")
	Scope    string // Either a sheet name or "Workbook" for global scope
}

// GetDefinedNames returns all named ranges in the workbook
func (ef *ExcelFile) GetDefinedNames() []DefinedNameInfo {
	names := ef.file.GetDefinedName()
	result := make([]DefinedNameInfo, len(names))
	for i, name := range names {
		result[i] = DefinedNameInfo{
			Name:     name.Name,
			RefersTo: name.RefersTo,
			Scope:    name.Scope,
		}
	}
	return result
}
