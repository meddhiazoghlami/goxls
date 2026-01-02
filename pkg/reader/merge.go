package reader

import (
	"excel-lite/pkg/models"

	"github.com/xuri/excelize/v2"
)

// MergeProcessor handles merged cell detection and expansion
type MergeProcessor struct {
	config models.DetectionConfig
}

// NewMergeProcessor creates a new merge processor
func NewMergeProcessor(config models.DetectionConfig) *MergeProcessor {
	return &MergeProcessor{config: config}
}

// ParsedMergeRange represents a parsed merge region with 0-indexed coordinates
type ParsedMergeRange struct {
	StartRow int
	StartCol int
	EndRow   int
	EndCol   int
	Value    string
}

// ParseMergeCells converts MergeCellInfo to ParsedMergeRange with coordinates
func (mp *MergeProcessor) ParseMergeCells(merges []MergeCellInfo) ([]ParsedMergeRange, error) {
	result := make([]ParsedMergeRange, 0, len(merges))

	for _, m := range merges {
		startCol, startRow, err := excelize.CellNameToCoordinates(m.StartCell)
		if err != nil {
			return nil, err
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(m.EndCell)
		if err != nil {
			return nil, err
		}

		// Skip single-cell "merges"
		if startRow == endRow && startCol == endCol {
			continue
		}

		// Convert to 0-indexed
		result = append(result, ParsedMergeRange{
			StartRow: startRow - 1,
			StartCol: startCol - 1,
			EndRow:   endRow - 1,
			EndCol:   endCol - 1,
			Value:    m.Value,
		})
	}

	return result, nil
}

// ApplyMerges applies merge information to a cell grid
func (mp *MergeProcessor) ApplyMerges(grid [][]models.Cell, merges []ParsedMergeRange) {
	for _, merge := range merges {
		mp.applyMerge(grid, merge)
	}
}

// applyMerge applies a single merge to the grid
func (mp *MergeProcessor) applyMerge(grid [][]models.Cell, merge ParsedMergeRange) {
	if len(grid) == 0 {
		return
	}

	// Get the origin cell's type for consistent typing across merge
	var originCellType models.CellType = models.CellTypeString
	var originValue interface{} = merge.Value

	if merge.StartRow < len(grid) && merge.StartCol < len(grid[merge.StartRow]) {
		originCell := &grid[merge.StartRow][merge.StartCol]
		originCellType = originCell.Type
		originValue = originCell.Value
	}

	for row := merge.StartRow; row <= merge.EndRow; row++ {
		if row < 0 || row >= len(grid) {
			continue
		}
		for col := merge.StartCol; col <= merge.EndCol; col++ {
			if col < 0 || col >= len(grid[row]) {
				continue
			}

			isOrigin := (row == merge.StartRow && col == merge.StartCol)
			cell := &grid[row][col]

			if mp.config.ExpandMergedCells {
				// Copy value from origin to all cells in the merge range
				cell.RawValue = merge.Value
				cell.Value = originValue
				if !isOrigin {
					cell.Type = originCellType
				}
			}

			if mp.config.TrackMergeMetadata {
				cell.IsMerged = true
				cell.MergeRange = &models.MergeRange{
					StartRow: merge.StartRow,
					StartCol: merge.StartCol,
					EndRow:   merge.EndRow,
					EndCol:   merge.EndCol,
					IsOrigin: isOrigin,
				}
			}
		}
	}
}

// BuildMergeMap creates a map for quick merge lookup by cell coordinates
func (mp *MergeProcessor) BuildMergeMap(merges []ParsedMergeRange) map[int]map[int]*ParsedMergeRange {
	result := make(map[int]map[int]*ParsedMergeRange)

	for i := range merges {
		merge := &merges[i]
		for row := merge.StartRow; row <= merge.EndRow; row++ {
			if result[row] == nil {
				result[row] = make(map[int]*ParsedMergeRange)
			}
			for col := merge.StartCol; col <= merge.EndCol; col++ {
				result[row][col] = merge
			}
		}
	}

	return result
}

// IsCellInMerge checks if a cell at given coordinates is part of any merge
func (mp *MergeProcessor) IsCellInMerge(merges []ParsedMergeRange, row, col int) *ParsedMergeRange {
	for i := range merges {
		merge := &merges[i]
		if row >= merge.StartRow && row <= merge.EndRow &&
			col >= merge.StartCol && col <= merge.EndCol {
			return merge
		}
	}
	return nil
}
