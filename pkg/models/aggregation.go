package models

import (
	"fmt"
	"math"
	"sort"
	"strings"
)

// AggregateOp represents the type of aggregation operation
type AggregateOp int

const (
	AggSum AggregateOp = iota
	AggCount
	AggAvg
	AggMin
	AggMax
)

// String returns the string representation of the aggregation operation
func (op AggregateOp) String() string {
	switch op {
	case AggSum:
		return "Sum"
	case AggCount:
		return "Count"
	case AggAvg:
		return "Avg"
	case AggMin:
		return "Min"
	case AggMax:
		return "Max"
	default:
		return "Unknown"
	}
}

// AggregateFunc defines an aggregation operation on a column
type AggregateFunc struct {
	Column string
	Op     AggregateOp
	Alias  string // Output column name (optional)
}

// OutputName returns the column name for the aggregation result
func (a AggregateFunc) OutputName() string {
	if a.Alias != "" {
		return a.Alias
	}
	return fmt.Sprintf("%s_%s", a.Op.String(), a.Column)
}

// As sets a custom alias for the output column name
func (a AggregateFunc) As(alias string) AggregateFunc {
	a.Alias = alias
	return a
}

// Sum creates a sum aggregation for the specified column
func Sum(column string) AggregateFunc {
	return AggregateFunc{Column: column, Op: AggSum}
}

// Count creates a count aggregation for the specified column
func Count(column string) AggregateFunc {
	return AggregateFunc{Column: column, Op: AggCount}
}

// Avg creates an average aggregation for the specified column
func Avg(column string) AggregateFunc {
	return AggregateFunc{Column: column, Op: AggAvg}
}

// Min creates a minimum aggregation for the specified column
func Min(column string) AggregateFunc {
	return AggregateFunc{Column: column, Op: AggMin}
}

// Max creates a maximum aggregation for the specified column
func Max(column string) AggregateFunc {
	return AggregateFunc{Column: column, Op: AggMax}
}

// GroupedTable represents a table grouped by one or more columns
type GroupedTable struct {
	source    *Table
	groupCols []string
}

// GroupBy groups the table by the specified columns and returns a GroupedTable
// for further aggregation operations.
func (t *Table) GroupBy(columns ...string) *GroupedTable {
	return &GroupedTable{
		source:    t,
		groupCols: columns,
	}
}

// groupKey generates a unique key for a row based on group column values
func (g *GroupedTable) groupKey(row Row) string {
	parts := make([]string, len(g.groupCols))
	for i, col := range g.groupCols {
		if cell, ok := row.Get(col); ok {
			parts[i] = cell.RawValue
		} else {
			parts[i] = ""
		}
	}
	return strings.Join(parts, "\x00") // Use null byte as separator
}

// aggregator holds the state for computing an aggregation
type aggregator struct {
	sum   float64
	count int
	min   float64
	max   float64
	hasValue bool
}

// Aggregate computes the specified aggregations for each group and returns
// a new table with the results. The resulting table has columns for each
// group column followed by columns for each aggregation result.
func (g *GroupedTable) Aggregate(funcs ...AggregateFunc) *Table {
	if g.source == nil || len(funcs) == 0 {
		return &Table{
			Headers: []string{},
			Rows:    []Row{},
		}
	}

	// Build headers: group columns + aggregation results
	headers := make([]string, 0, len(g.groupCols)+len(funcs))
	headers = append(headers, g.groupCols...)
	for _, f := range funcs {
		headers = append(headers, f.OutputName())
	}

	// Group rows by key
	groups := make(map[string][]Row)
	groupOrder := make([]string, 0) // Preserve insertion order

	for _, row := range g.source.Rows {
		key := g.groupKey(row)
		if _, exists := groups[key]; !exists {
			groupOrder = append(groupOrder, key)
		}
		groups[key] = append(groups[key], row)
	}

	// Sort groups for deterministic output
	sort.Strings(groupOrder)

	// Compute aggregations for each group
	resultRows := make([]Row, 0, len(groups))

	for _, key := range groupOrder {
		rows := groups[key]
		if len(rows) == 0 {
			continue
		}

		// Initialize aggregators for each function
		aggs := make([]aggregator, len(funcs))
		for i := range aggs {
			aggs[i].min = math.MaxFloat64
			aggs[i].max = -math.MaxFloat64
		}

		// Process each row in the group
		for _, row := range rows {
			for i, f := range funcs {
				cell, ok := row.Get(f.Column)
				if !ok {
					continue
				}

				switch f.Op {
				case AggCount:
					if !cell.IsEmpty() {
						aggs[i].count++
					}
				case AggSum, AggAvg:
					if val, ok := cell.AsFloat(); ok {
						aggs[i].sum += val
						aggs[i].count++
						aggs[i].hasValue = true
					}
				case AggMin:
					if val, ok := cell.AsFloat(); ok {
						if val < aggs[i].min {
							aggs[i].min = val
						}
						aggs[i].hasValue = true
					}
				case AggMax:
					if val, ok := cell.AsFloat(); ok {
						if val > aggs[i].max {
							aggs[i].max = val
						}
						aggs[i].hasValue = true
					}
				}
			}
		}

		// Build result row
		resultRow := Row{
			Index:  len(resultRows),
			Values: make(map[string]Cell),
			Cells:  make([]Cell, 0, len(headers)),
		}

		// Add group column values (from first row in group)
		firstRow := rows[0]
		for _, col := range g.groupCols {
			if cell, ok := firstRow.Get(col); ok {
				resultRow.Values[col] = cell
				resultRow.Cells = append(resultRow.Cells, cell)
			} else {
				emptyCell := Cell{Type: CellTypeEmpty}
				resultRow.Values[col] = emptyCell
				resultRow.Cells = append(resultRow.Cells, emptyCell)
			}
		}

		// Add aggregation results
		for i, f := range funcs {
			var cell Cell
			outputName := f.OutputName()

			switch f.Op {
			case AggCount:
				cell = Cell{
					Value:    float64(aggs[i].count),
					Type:     CellTypeNumber,
					RawValue: fmt.Sprintf("%d", aggs[i].count),
				}
			case AggSum:
				if aggs[i].hasValue {
					cell = Cell{
						Value:    aggs[i].sum,
						Type:     CellTypeNumber,
						RawValue: formatFloat(aggs[i].sum),
					}
				} else {
					cell = Cell{Type: CellTypeEmpty, RawValue: ""}
				}
			case AggAvg:
				if aggs[i].count > 0 && aggs[i].hasValue {
					avg := aggs[i].sum / float64(aggs[i].count)
					cell = Cell{
						Value:    avg,
						Type:     CellTypeNumber,
						RawValue: formatFloat(avg),
					}
				} else {
					cell = Cell{Type: CellTypeEmpty, RawValue: ""}
				}
			case AggMin:
				if aggs[i].hasValue {
					cell = Cell{
						Value:    aggs[i].min,
						Type:     CellTypeNumber,
						RawValue: formatFloat(aggs[i].min),
					}
				} else {
					cell = Cell{Type: CellTypeEmpty, RawValue: ""}
				}
			case AggMax:
				if aggs[i].hasValue {
					cell = Cell{
						Value:    aggs[i].max,
						Type:     CellTypeNumber,
						RawValue: formatFloat(aggs[i].max),
					}
				} else {
					cell = Cell{Type: CellTypeEmpty, RawValue: ""}
				}
			}

			resultRow.Values[outputName] = cell
			resultRow.Cells = append(resultRow.Cells, cell)
		}

		resultRows = append(resultRows, resultRow)
	}

	return &Table{
		Name:    g.source.Name,
		Headers: headers,
		Rows:    resultRows,
	}
}

// formatFloat formats a float64 for display, removing unnecessary trailing zeros
func formatFloat(f float64) string {
	s := fmt.Sprintf("%f", f)
	// Trim trailing zeros after decimal point
	if strings.Contains(s, ".") {
		s = strings.TrimRight(s, "0")
		s = strings.TrimRight(s, ".")
	}
	return s
}
