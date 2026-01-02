package models

import (
	"testing"
	"time"
)

// =============================================================================
// Cell Tests
// =============================================================================

func TestCell_IsEmpty(t *testing.T) {
	tests := []struct {
		name     string
		cell     Cell
		expected bool
	}{
		{
			name:     "empty cell type",
			cell:     Cell{Type: CellTypeEmpty, RawValue: ""},
			expected: true,
		},
		{
			name:     "empty raw value",
			cell:     Cell{Type: CellTypeString, RawValue: ""},
			expected: true,
		},
		{
			name:     "non-empty string",
			cell:     Cell{Type: CellTypeString, RawValue: "hello"},
			expected: false,
		},
		{
			name:     "number cell",
			cell:     Cell{Type: CellTypeNumber, RawValue: "42"},
			expected: false,
		},
		{
			name:     "whitespace only",
			cell:     Cell{Type: CellTypeString, RawValue: "   "},
			expected: false,
		},
		{
			name:     "zero value",
			cell:     Cell{Type: CellTypeNumber, RawValue: "0"},
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.cell.IsEmpty(); got != tt.expected {
				t.Errorf("Cell.IsEmpty() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestCell_AsString(t *testing.T) {
	tests := []struct {
		name     string
		cell     Cell
		expected string
	}{
		{
			name:     "nil value",
			cell:     Cell{Value: nil, RawValue: ""},
			expected: "",
		},
		{
			name:     "string value",
			cell:     Cell{Value: "hello", RawValue: "hello"},
			expected: "hello",
		},
		{
			name:     "number value returns raw",
			cell:     Cell{Value: 42.5, RawValue: "42.5"},
			expected: "42.5",
		},
		{
			name:     "bool value returns raw",
			cell:     Cell{Value: true, RawValue: "true"},
			expected: "true",
		},
		{
			name:     "empty string",
			cell:     Cell{Value: "", RawValue: ""},
			expected: "",
		},
		{
			name:     "unicode string",
			cell:     Cell{Value: "日本語", RawValue: "日本語"},
			expected: "日本語",
		},
		{
			name:     "string with special chars",
			cell:     Cell{Value: "hello\nworld\ttab", RawValue: "hello\nworld\ttab"},
			expected: "hello\nworld\ttab",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.cell.AsString(); got != tt.expected {
				t.Errorf("Cell.AsString() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestCell_AsFloat(t *testing.T) {
	tests := []struct {
		name        string
		cell        Cell
		expected    float64
		expectedOk  bool
	}{
		{
			name:       "valid float",
			cell:       Cell{Value: 42.5},
			expected:   42.5,
			expectedOk: true,
		},
		{
			name:       "zero",
			cell:       Cell{Value: 0.0},
			expected:   0.0,
			expectedOk: true,
		},
		{
			name:       "negative float",
			cell:       Cell{Value: -123.456},
			expected:   -123.456,
			expectedOk: true,
		},
		{
			name:       "string value",
			cell:       Cell{Value: "not a number"},
			expected:   0,
			expectedOk: false,
		},
		{
			name:       "nil value",
			cell:       Cell{Value: nil},
			expected:   0,
			expectedOk: false,
		},
		{
			name:       "int value (not float64)",
			cell:       Cell{Value: 42},
			expected:   0,
			expectedOk: false,
		},
		{
			name:       "very large float",
			cell:       Cell{Value: 1e308},
			expected:   1e308,
			expectedOk: true,
		},
		{
			name:       "very small float",
			cell:       Cell{Value: 1e-308},
			expected:   1e-308,
			expectedOk: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, ok := tt.cell.AsFloat()
			if ok != tt.expectedOk {
				t.Errorf("Cell.AsFloat() ok = %v, want %v", ok, tt.expectedOk)
			}
			if got != tt.expected {
				t.Errorf("Cell.AsFloat() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestCell_AsTime(t *testing.T) {
	now := time.Now()
	fixedTime := time.Date(2024, 1, 15, 10, 30, 0, 0, time.UTC)

	tests := []struct {
		name       string
		cell       Cell
		expected   time.Time
		expectedOk bool
	}{
		{
			name:       "valid time",
			cell:       Cell{Value: fixedTime},
			expected:   fixedTime,
			expectedOk: true,
		},
		{
			name:       "current time",
			cell:       Cell{Value: now},
			expected:   now,
			expectedOk: true,
		},
		{
			name:       "string value",
			cell:       Cell{Value: "2024-01-15"},
			expected:   time.Time{},
			expectedOk: false,
		},
		{
			name:       "nil value",
			cell:       Cell{Value: nil},
			expected:   time.Time{},
			expectedOk: false,
		},
		{
			name:       "number value",
			cell:       Cell{Value: 45678.0},
			expected:   time.Time{},
			expectedOk: false,
		},
		{
			name:       "zero time",
			cell:       Cell{Value: time.Time{}},
			expected:   time.Time{},
			expectedOk: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, ok := tt.cell.AsTime()
			if ok != tt.expectedOk {
				t.Errorf("Cell.AsTime() ok = %v, want %v", ok, tt.expectedOk)
			}
			if !got.Equal(tt.expected) {
				t.Errorf("Cell.AsTime() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestCell_IsMergeOrigin(t *testing.T) {
	tests := []struct {
		name     string
		cell     Cell
		expected bool
	}{
		{
			name:     "nil MergeRange",
			cell:     Cell{Value: "test", Type: CellTypeString},
			expected: false,
		},
		{
			name: "MergeRange with IsOrigin true",
			cell: Cell{
				Value:    "merged",
				Type:     CellTypeString,
				IsMerged: true,
				MergeRange: &MergeRange{
					StartRow: 0,
					StartCol: 0,
					EndRow:   1,
					EndCol:   1,
					IsOrigin: true,
				},
			},
			expected: true,
		},
		{
			name: "MergeRange with IsOrigin false",
			cell: Cell{
				Value:    "merged",
				Type:     CellTypeString,
				IsMerged: true,
				MergeRange: &MergeRange{
					StartRow: 0,
					StartCol: 0,
					EndRow:   1,
					EndCol:   1,
					IsOrigin: false,
				},
			},
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := tt.cell.IsMergeOrigin()
			if got != tt.expected {
				t.Errorf("Cell.IsMergeOrigin() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// =============================================================================
// Row Tests
// =============================================================================

func TestRow_Get(t *testing.T) {
	row := Row{
		Index: 1,
		Values: map[string]Cell{
			"Name":  {Value: "Alice", Type: CellTypeString, RawValue: "Alice"},
			"Age":   {Value: 30.0, Type: CellTypeNumber, RawValue: "30"},
			"Email": {Value: "alice@test.com", Type: CellTypeString, RawValue: "alice@test.com"},
		},
	}

	tests := []struct {
		name       string
		header     string
		expectedOk bool
		checkValue string
	}{
		{
			name:       "existing header",
			header:     "Name",
			expectedOk: true,
			checkValue: "Alice",
		},
		{
			name:       "another existing header",
			header:     "Age",
			expectedOk: true,
			checkValue: "30",
		},
		{
			name:       "non-existing header",
			header:     "Phone",
			expectedOk: false,
		},
		{
			name:       "empty header",
			header:     "",
			expectedOk: false,
		},
		{
			name:       "case sensitive - wrong case",
			header:     "name",
			expectedOk: false,
		},
		{
			name:       "header with spaces",
			header:     "Name ",
			expectedOk: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cell, ok := row.Get(tt.header)
			if ok != tt.expectedOk {
				t.Errorf("Row.Get(%q) ok = %v, want %v", tt.header, ok, tt.expectedOk)
			}
			if ok && cell.AsString() != tt.checkValue {
				t.Errorf("Row.Get(%q) value = %v, want %v", tt.header, cell.AsString(), tt.checkValue)
			}
		})
	}
}

func TestRow_Get_EmptyRow(t *testing.T) {
	row := Row{
		Index:  0,
		Values: map[string]Cell{},
	}

	cell, ok := row.Get("anything")
	if ok {
		t.Error("Expected ok to be false for empty row")
	}
	if cell.Value != nil {
		t.Error("Expected nil value for missing cell")
	}
}

// =============================================================================
// Table Tests
// =============================================================================

func TestTable_RowCount(t *testing.T) {
	tests := []struct {
		name     string
		table    Table
		expected int
	}{
		{
			name:     "empty table",
			table:    Table{Rows: []Row{}},
			expected: 0,
		},
		{
			name:     "single row",
			table:    Table{Rows: []Row{{Index: 1}}},
			expected: 1,
		},
		{
			name: "multiple rows",
			table: Table{Rows: []Row{
				{Index: 1},
				{Index: 2},
				{Index: 3},
				{Index: 4},
				{Index: 5},
			}},
			expected: 5,
		},
		{
			name:     "nil rows",
			table:    Table{},
			expected: 0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.table.RowCount(); got != tt.expected {
				t.Errorf("Table.RowCount() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestTable_ColCount(t *testing.T) {
	tests := []struct {
		name     string
		table    Table
		expected int
	}{
		{
			name:     "empty headers",
			table:    Table{Headers: []string{}},
			expected: 0,
		},
		{
			name:     "single column",
			table:    Table{Headers: []string{"ID"}},
			expected: 1,
		},
		{
			name:     "multiple columns",
			table:    Table{Headers: []string{"ID", "Name", "Email", "Phone", "Address"}},
			expected: 5,
		},
		{
			name:     "nil headers",
			table:    Table{},
			expected: 0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.table.ColCount(); got != tt.expected {
				t.Errorf("Table.ColCount() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// =============================================================================
// Filter Tests
// =============================================================================

func TestTable_Filter_ByNumericValue(t *testing.T) {
	table := Table{
		Name:      "TestTable",
		Headers:   []string{"Name", "Age"},
		HeaderRow: 0,
		StartRow:  0,
		EndRow:    4,
		StartCol:  0,
		EndCol:    1,
		Rows: []Row{
			{Index: 1, Values: map[string]Cell{"Name": {Value: "Alice", RawValue: "Alice"}, "Age": {Value: 30.0, Type: CellTypeNumber, RawValue: "30"}}},
			{Index: 2, Values: map[string]Cell{"Name": {Value: "Bob", RawValue: "Bob"}, "Age": {Value: 17.0, Type: CellTypeNumber, RawValue: "17"}}},
			{Index: 3, Values: map[string]Cell{"Name": {Value: "Charlie", RawValue: "Charlie"}, "Age": {Value: 25.0, Type: CellTypeNumber, RawValue: "25"}}},
			{Index: 4, Values: map[string]Cell{"Name": {Value: "Diana", RawValue: "Diana"}, "Age": {Value: 16.0, Type: CellTypeNumber, RawValue: "16"}}},
		},
	}

	// Filter rows where Age > 18
	filtered := table.Filter(func(row Row) bool {
		if cell, ok := row.Get("Age"); ok {
			if val, ok := cell.AsFloat(); ok {
				return val > 18
			}
		}
		return false
	})

	if filtered.RowCount() != 2 {
		t.Errorf("Filtered table should have 2 rows, got %d", filtered.RowCount())
	}

	// Verify the correct rows were kept
	names := make([]string, 0)
	for _, row := range filtered.Rows {
		if cell, ok := row.Get("Name"); ok {
			names = append(names, cell.AsString())
		}
	}
	if names[0] != "Alice" || names[1] != "Charlie" {
		t.Errorf("Expected [Alice, Charlie], got %v", names)
	}
}

func TestTable_Filter_ByStringValue(t *testing.T) {
	table := Table{
		Name:    "TestTable",
		Headers: []string{"Status"},
		Rows: []Row{
			{Values: map[string]Cell{"Status": {Value: "active", RawValue: "active"}}},
			{Values: map[string]Cell{"Status": {Value: "inactive", RawValue: "inactive"}}},
			{Values: map[string]Cell{"Status": {Value: "active", RawValue: "active"}}},
			{Values: map[string]Cell{"Status": {Value: "pending", RawValue: "pending"}}},
		},
	}

	filtered := table.Filter(func(row Row) bool {
		if cell, ok := row.Get("Status"); ok {
			return cell.AsString() == "active"
		}
		return false
	})

	if filtered.RowCount() != 2 {
		t.Errorf("Filtered table should have 2 active rows, got %d", filtered.RowCount())
	}
}

func TestTable_Filter_NoMatches(t *testing.T) {
	table := Table{
		Name:    "TestTable",
		Headers: []string{"Value"},
		Rows: []Row{
			{Values: map[string]Cell{"Value": {Value: 1.0, Type: CellTypeNumber}}},
			{Values: map[string]Cell{"Value": {Value: 2.0, Type: CellTypeNumber}}},
			{Values: map[string]Cell{"Value": {Value: 3.0, Type: CellTypeNumber}}},
		},
	}

	filtered := table.Filter(func(row Row) bool {
		if cell, ok := row.Get("Value"); ok {
			if val, ok := cell.AsFloat(); ok {
				return val > 100 // No values match
			}
		}
		return false
	})

	if filtered.RowCount() != 0 {
		t.Errorf("Filtered table should be empty, got %d rows", filtered.RowCount())
	}
}

func TestTable_Filter_AllMatch(t *testing.T) {
	table := Table{
		Name:    "TestTable",
		Headers: []string{"Value"},
		Rows: []Row{
			{Values: map[string]Cell{"Value": {Value: 10.0, Type: CellTypeNumber}}},
			{Values: map[string]Cell{"Value": {Value: 20.0, Type: CellTypeNumber}}},
			{Values: map[string]Cell{"Value": {Value: 30.0, Type: CellTypeNumber}}},
		},
	}

	filtered := table.Filter(func(row Row) bool {
		return true // All rows match
	})

	if filtered.RowCount() != 3 {
		t.Errorf("All rows should match, got %d", filtered.RowCount())
	}
}

func TestTable_Filter_EmptyTable(t *testing.T) {
	table := Table{
		Name:    "EmptyTable",
		Headers: []string{"Col1"},
		Rows:    []Row{},
	}

	filtered := table.Filter(func(row Row) bool {
		return true
	})

	if filtered.RowCount() != 0 {
		t.Errorf("Filtered empty table should remain empty")
	}
	if filtered.Name != "EmptyTable" {
		t.Errorf("Table name should be preserved")
	}
}

func TestTable_Filter_PreservesMetadata(t *testing.T) {
	table := Table{
		Name:      "OriginalTable",
		Headers:   []string{"A", "B", "C"},
		HeaderRow: 5,
		StartRow:  5,
		EndRow:    15,
		StartCol:  2,
		EndCol:    4,
		Rows: []Row{
			{Values: map[string]Cell{"A": {Value: 1.0}}},
			{Values: map[string]Cell{"A": {Value: 2.0}}},
		},
	}

	filtered := table.Filter(func(row Row) bool {
		return true
	})

	if filtered.Name != table.Name {
		t.Errorf("Name not preserved: got %s, want %s", filtered.Name, table.Name)
	}
	if len(filtered.Headers) != len(table.Headers) {
		t.Errorf("Headers not preserved")
	}
	if filtered.HeaderRow != table.HeaderRow {
		t.Errorf("HeaderRow not preserved")
	}
	if filtered.StartRow != table.StartRow {
		t.Errorf("StartRow not preserved")
	}
	if filtered.EndRow != table.EndRow {
		t.Errorf("EndRow not preserved")
	}
	if filtered.StartCol != table.StartCol {
		t.Errorf("StartCol not preserved")
	}
	if filtered.EndCol != table.EndCol {
		t.Errorf("EndCol not preserved")
	}
}

func TestTable_Filter_Chaining(t *testing.T) {
	table := Table{
		Name:    "TestTable",
		Headers: []string{"Name", "Age", "Status"},
		Rows: []Row{
			{Values: map[string]Cell{
				"Name":   {Value: "Alice", RawValue: "Alice"},
				"Age":    {Value: 30.0, Type: CellTypeNumber},
				"Status": {Value: "active", RawValue: "active"},
			}},
			{Values: map[string]Cell{
				"Name":   {Value: "Bob", RawValue: "Bob"},
				"Age":    {Value: 25.0, Type: CellTypeNumber},
				"Status": {Value: "inactive", RawValue: "inactive"},
			}},
			{Values: map[string]Cell{
				"Name":   {Value: "Charlie", RawValue: "Charlie"},
				"Age":    {Value: 35.0, Type: CellTypeNumber},
				"Status": {Value: "active", RawValue: "active"},
			}},
			{Values: map[string]Cell{
				"Name":   {Value: "Diana", RawValue: "Diana"},
				"Age":    {Value: 28.0, Type: CellTypeNumber},
				"Status": {Value: "active", RawValue: "active"},
			}},
		},
	}

	// Chain filters: active users over 30
	filtered := table.Filter(func(row Row) bool {
		if cell, ok := row.Get("Status"); ok {
			return cell.AsString() == "active"
		}
		return false
	}).Filter(func(row Row) bool {
		if cell, ok := row.Get("Age"); ok {
			if val, ok := cell.AsFloat(); ok {
				return val > 30
			}
		}
		return false
	})

	// Only Charlie (35) is active and over 30
	if filtered.RowCount() != 1 {
		t.Errorf("Chained filter should have 1 row (Charlie), got %d", filtered.RowCount())
	}
}

func TestTable_Filter_DoesNotModifyOriginal(t *testing.T) {
	table := Table{
		Name:    "Original",
		Headers: []string{"Value"},
		Rows: []Row{
			{Values: map[string]Cell{"Value": {Value: 1.0}}},
			{Values: map[string]Cell{"Value": {Value: 2.0}}},
			{Values: map[string]Cell{"Value": {Value: 3.0}}},
		},
	}

	originalRowCount := table.RowCount()

	_ = table.Filter(func(row Row) bool {
		return false // Filter out all
	})

	if table.RowCount() != originalRowCount {
		t.Errorf("Original table was modified: got %d rows, want %d", table.RowCount(), originalRowCount)
	}
}

// =============================================================================
// Deduplication Tests
// =============================================================================

func TestTable_FindDuplicates(t *testing.T) {
	table := Table{
		Headers: []string{"Email", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"Email": {RawValue: "bob@test.com"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice Duplicate"}}},
			{Values: map[string]Cell{"Email": {RawValue: "charlie@test.com"}, "Name": {RawValue: "Charlie"}}},
			{Values: map[string]Cell{"Email": {RawValue: "bob@test.com"}, "Name": {RawValue: "Bob Duplicate"}}},
		},
	}

	duplicates := table.FindDuplicates("Email")

	if len(duplicates) != 2 {
		t.Errorf("Expected 2 duplicates, got %d", len(duplicates))
	}
}

func TestTable_FindDuplicates_NoDuplicates(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	duplicates := table.FindDuplicates("ID")

	if len(duplicates) != 0 {
		t.Errorf("Expected 0 duplicates, got %d", len(duplicates))
	}
}

func TestTable_FindDuplicates_AllDuplicates(t *testing.T) {
	table := Table{
		Headers: []string{"Status", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"Status": {RawValue: "active"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"Status": {RawValue: "active"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"Status": {RawValue: "active"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	duplicates := table.FindDuplicates("Status")

	// First occurrence is not a duplicate, so 2 duplicates
	if len(duplicates) != 2 {
		t.Errorf("Expected 2 duplicates, got %d", len(duplicates))
	}
}

func TestTable_FindDuplicates_EmptyTable(t *testing.T) {
	table := Table{
		Headers: []string{"ID"},
		Rows:    []Row{},
	}

	duplicates := table.FindDuplicates("ID")

	if len(duplicates) != 0 {
		t.Errorf("Expected 0 duplicates for empty table, got %d", len(duplicates))
	}
}

func TestTable_Deduplicate(t *testing.T) {
	table := Table{
		Name:    "TestTable",
		Headers: []string{"Email", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"Email": {RawValue: "bob@test.com"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice Duplicate"}}},
			{Values: map[string]Cell{"Email": {RawValue: "charlie@test.com"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	deduped := table.Deduplicate("Email")

	if deduped.RowCount() != 3 {
		t.Errorf("Expected 3 rows after dedup, got %d", deduped.RowCount())
	}

	// Verify first occurrence is kept
	if cell, ok := deduped.Rows[0].Get("Name"); ok {
		if cell.RawValue != "Alice" {
			t.Errorf("Expected first Alice to be kept, got %s", cell.RawValue)
		}
	}
}

func TestTable_Deduplicate_PreservesMetadata(t *testing.T) {
	table := Table{
		Name:      "OriginalTable",
		Headers:   []string{"ID", "Name"},
		HeaderRow: 5,
		StartRow:  5,
		EndRow:    10,
		StartCol:  2,
		EndCol:    3,
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
		},
	}

	deduped := table.Deduplicate("ID")

	if deduped.Name != table.Name {
		t.Errorf("Name not preserved")
	}
	if deduped.HeaderRow != table.HeaderRow {
		t.Errorf("HeaderRow not preserved")
	}
	if deduped.StartRow != table.StartRow {
		t.Errorf("StartRow not preserved")
	}
}

func TestTable_Deduplicate_NoDuplicates(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	deduped := table.Deduplicate("ID")

	if deduped.RowCount() != 3 {
		t.Errorf("Expected 3 rows (no change), got %d", deduped.RowCount())
	}
}

func TestTable_Deduplicate_DoesNotModifyOriginal(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice Dup"}}},
		},
	}

	originalCount := table.RowCount()
	_ = table.Deduplicate("ID")

	if table.RowCount() != originalCount {
		t.Errorf("Original table was modified")
	}
}

func TestTable_FindDuplicateGroups(t *testing.T) {
	table := Table{
		Headers: []string{"Email", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice 1"}}},
			{Values: map[string]Cell{"Email": {RawValue: "bob@test.com"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice 2"}}},
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice 3"}}},
		},
	}

	groups := table.FindDuplicateGroups("Email")

	if len(groups) != 1 {
		t.Errorf("Expected 1 duplicate group, got %d", len(groups))
	}

	if groups[0].Count != 3 {
		t.Errorf("Expected group count of 3, got %d", groups[0].Count)
	}

	if groups[0].KeyValue != "alice@test.com" {
		t.Errorf("Expected key 'alice@test.com', got '%s'", groups[0].KeyValue)
	}
}

func TestTable_FindDuplicateGroups_MultipleDuplicates(t *testing.T) {
	table := Table{
		Headers: []string{"Status", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"Status": {RawValue: "active"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"Status": {RawValue: "inactive"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"Status": {RawValue: "active"}, "Name": {RawValue: "Charlie"}}},
			{Values: map[string]Cell{"Status": {RawValue: "inactive"}, "Name": {RawValue: "Diana"}}},
			{Values: map[string]Cell{"Status": {RawValue: "pending"}, "Name": {RawValue: "Eve"}}},
		},
	}

	groups := table.FindDuplicateGroups("Status")

	if len(groups) != 2 {
		t.Errorf("Expected 2 duplicate groups (active, inactive), got %d", len(groups))
	}
}

func TestTable_FindDuplicateGroups_NoDuplicates(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	groups := table.FindDuplicateGroups("ID")

	if len(groups) != 0 {
		t.Errorf("Expected 0 duplicate groups, got %d", len(groups))
	}
}

// =============================================================================
// Column Transformation Tests
// =============================================================================

func TestTable_Select(t *testing.T) {
	table := Table{
		Name:    "TestTable",
		Headers: []string{"ID", "Name", "Email", "Age"},
		Rows: []Row{
			{Values: map[string]Cell{
				"ID":    {RawValue: "1"},
				"Name":  {RawValue: "Alice"},
				"Email": {RawValue: "alice@test.com"},
				"Age":   {RawValue: "30"},
			}},
			{Values: map[string]Cell{
				"ID":    {RawValue: "2"},
				"Name":  {RawValue: "Bob"},
				"Email": {RawValue: "bob@test.com"},
				"Age":   {RawValue: "25"},
			}},
		},
	}

	selected := table.Select("Name", "Email")

	if len(selected.Headers) != 2 {
		t.Errorf("Expected 2 headers, got %d", len(selected.Headers))
	}
	if selected.Headers[0] != "Name" || selected.Headers[1] != "Email" {
		t.Errorf("Expected headers [Name, Email], got %v", selected.Headers)
	}

	// Verify row data
	if cell, ok := selected.Rows[0].Get("Name"); !ok || cell.RawValue != "Alice" {
		t.Errorf("Expected Name=Alice in first row")
	}
	if _, ok := selected.Rows[0].Get("ID"); ok {
		t.Errorf("ID column should not be present")
	}
}

func TestTable_Select_InvalidColumn(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
		},
	}

	selected := table.Select("Name", "NonExistent")

	if len(selected.Headers) != 1 {
		t.Errorf("Expected 1 header (invalid column ignored), got %d", len(selected.Headers))
	}
	if selected.Headers[0] != "Name" {
		t.Errorf("Expected header [Name], got %v", selected.Headers)
	}
}

func TestTable_Select_EmptyTable(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name"},
		Rows:    []Row{},
	}

	selected := table.Select("Name")

	if len(selected.Headers) != 1 {
		t.Errorf("Expected 1 header, got %d", len(selected.Headers))
	}
	if len(selected.Rows) != 0 {
		t.Errorf("Expected 0 rows, got %d", len(selected.Rows))
	}
}

func TestTable_Select_PreservesMetadata(t *testing.T) {
	table := Table{
		Name:      "Original",
		Headers:   []string{"A", "B"},
		HeaderRow: 5,
		StartRow:  5,
		EndRow:    10,
		Rows:      []Row{},
	}

	selected := table.Select("A")

	if selected.Name != "Original" {
		t.Errorf("Name not preserved")
	}
	if selected.HeaderRow != 5 {
		t.Errorf("HeaderRow not preserved")
	}
}

func TestTable_Rename(t *testing.T) {
	table := Table{
		Headers: []string{"old_name", "old_email"},
		Rows: []Row{
			{Values: map[string]Cell{
				"old_name":  {RawValue: "Alice"},
				"old_email": {RawValue: "alice@test.com"},
			}},
		},
	}

	renamed := table.Rename(map[string]string{
		"old_name":  "Name",
		"old_email": "Email",
	})

	if renamed.Headers[0] != "Name" || renamed.Headers[1] != "Email" {
		t.Errorf("Expected headers [Name, Email], got %v", renamed.Headers)
	}

	// Verify row values accessible by new names
	if cell, ok := renamed.Rows[0].Get("Name"); !ok || cell.RawValue != "Alice" {
		t.Errorf("Expected to get Alice by new name 'Name'")
	}
	if cell, ok := renamed.Rows[0].Get("Email"); !ok || cell.RawValue != "alice@test.com" {
		t.Errorf("Expected to get email by new name 'Email'")
	}
}

func TestTable_Rename_PartialMapping(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "old_name", "Status"},
		Rows: []Row{
			{Values: map[string]Cell{
				"ID":       {RawValue: "1"},
				"old_name": {RawValue: "Alice"},
				"Status":   {RawValue: "active"},
			}},
		},
	}

	renamed := table.Rename(map[string]string{
		"old_name": "Name",
	})

	// ID and Status should remain unchanged
	if renamed.Headers[0] != "ID" {
		t.Errorf("Expected ID to remain unchanged")
	}
	if renamed.Headers[1] != "Name" {
		t.Errorf("Expected old_name to be renamed to Name")
	}
	if renamed.Headers[2] != "Status" {
		t.Errorf("Expected Status to remain unchanged")
	}
}

func TestTable_Rename_DoesNotModifyOriginal(t *testing.T) {
	table := Table{
		Headers: []string{"old_name"},
		Rows: []Row{
			{Values: map[string]Cell{"old_name": {RawValue: "Alice"}}},
		},
	}

	_ = table.Rename(map[string]string{"old_name": "Name"})

	if table.Headers[0] != "old_name" {
		t.Errorf("Original table was modified")
	}
}

func TestTable_Reorder(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name", "Email", "Age"},
		Rows: []Row{
			{Values: map[string]Cell{
				"ID":    {RawValue: "1"},
				"Name":  {RawValue: "Alice"},
				"Email": {RawValue: "alice@test.com"},
				"Age":   {RawValue: "30"},
			}},
		},
	}

	reordered := table.Reorder("Email", "Name", "Age", "ID")

	expected := []string{"Email", "Name", "Age", "ID"}
	for i, h := range expected {
		if reordered.Headers[i] != h {
			t.Errorf("Expected header %d to be %s, got %s", i, h, reordered.Headers[i])
		}
	}
}

func TestTable_Reorder_Subset(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name", "Email", "Age"},
		Rows: []Row{
			{Values: map[string]Cell{
				"ID":    {RawValue: "1"},
				"Name":  {RawValue: "Alice"},
				"Email": {RawValue: "alice@test.com"},
				"Age":   {RawValue: "30"},
			}},
		},
	}

	// Reorder with only some columns - others are excluded
	reordered := table.Reorder("Email", "Name")

	if len(reordered.Headers) != 2 {
		t.Errorf("Expected 2 headers, got %d", len(reordered.Headers))
	}
	if reordered.Headers[0] != "Email" || reordered.Headers[1] != "Name" {
		t.Errorf("Expected [Email, Name], got %v", reordered.Headers)
	}
}

func TestTable_Reorder_InvalidColumn(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
		},
	}

	reordered := table.Reorder("Name", "NonExistent", "ID")

	// NonExistent should be ignored
	if len(reordered.Headers) != 2 {
		t.Errorf("Expected 2 headers (invalid ignored), got %d", len(reordered.Headers))
	}
	if reordered.Headers[0] != "Name" || reordered.Headers[1] != "ID" {
		t.Errorf("Expected [Name, ID], got %v", reordered.Headers)
	}
}

func TestTable_Select_DoesNotModifyOriginal(t *testing.T) {
	table := Table{
		Headers: []string{"ID", "Name", "Email"},
		Rows: []Row{
			{Values: map[string]Cell{
				"ID":    {RawValue: "1"},
				"Name":  {RawValue: "Alice"},
				"Email": {RawValue: "alice@test.com"},
			}},
		},
	}

	_ = table.Select("Name")

	if len(table.Headers) != 3 {
		t.Errorf("Original table was modified")
	}
}

func TestTable_Chaining_Transformations(t *testing.T) {
	table := Table{
		Headers: []string{"id", "user_name", "user_email", "status"},
		Rows: []Row{
			{Values: map[string]Cell{
				"id":         {RawValue: "1"},
				"user_name":  {RawValue: "Alice"},
				"user_email": {RawValue: "alice@test.com"},
				"status":     {RawValue: "active"},
			}},
		},
	}

	// Chain: Select -> Rename -> Reorder
	result := table.
		Select("user_name", "user_email", "status").
		Rename(map[string]string{"user_name": "Name", "user_email": "Email"}).
		Reorder("Email", "Name", "status")

	if len(result.Headers) != 3 {
		t.Errorf("Expected 3 headers, got %d", len(result.Headers))
	}
	if result.Headers[0] != "Email" {
		t.Errorf("Expected first header to be Email, got %s", result.Headers[0])
	}

	if cell, ok := result.Rows[0].Get("Email"); !ok || cell.RawValue != "alice@test.com" {
		t.Errorf("Expected Email value")
	}
}

// =============================================================================
// AnalyzeColumns Tests
// =============================================================================

func TestTable_AnalyzeColumns(t *testing.T) {
	table := Table{
		Headers: []string{"Name", "Age", "Active"},
		Rows: []Row{
			{
				Values: map[string]Cell{
					"Name":   {Value: "Alice", Type: CellTypeString, RawValue: "Alice"},
					"Age":    {Value: 30.0, Type: CellTypeNumber, RawValue: "30"},
					"Active": {Value: true, Type: CellTypeBool, RawValue: "true"},
				},
			},
			{
				Values: map[string]Cell{
					"Name":   {Value: "Bob", Type: CellTypeString, RawValue: "Bob"},
					"Age":    {Value: 25.0, Type: CellTypeNumber, RawValue: "25"},
					"Active": {Value: false, Type: CellTypeBool, RawValue: "false"},
				},
			},
			{
				Values: map[string]Cell{
					"Name":   {Value: "Charlie", Type: CellTypeString, RawValue: "Charlie"},
					"Age":    {Value: 35.0, Type: CellTypeNumber, RawValue: "35"},
					"Active": {Value: true, Type: CellTypeBool, RawValue: "true"},
				},
			},
		},
	}

	stats := table.AnalyzeColumns()

	if len(stats) != 3 {
		t.Fatalf("AnalyzeColumns() returned %d stats, want 3", len(stats))
	}

	// Check Name column
	if stats[0].Name != "Name" {
		t.Errorf("stats[0].Name = %q, want %q", stats[0].Name, "Name")
	}
	if stats[0].InferredType != CellTypeString {
		t.Errorf("stats[0].InferredType = %v, want CellTypeString", stats[0].InferredType)
	}
	if stats[0].StringCount != 3 {
		t.Errorf("stats[0].StringCount = %d, want 3", stats[0].StringCount)
	}
	if stats[0].UniqueCount != 3 {
		t.Errorf("stats[0].UniqueCount = %d, want 3", stats[0].UniqueCount)
	}

	// Check Age column
	if stats[1].InferredType != CellTypeNumber {
		t.Errorf("stats[1].InferredType = %v, want CellTypeNumber", stats[1].InferredType)
	}
	if stats[1].NumberCount != 3 {
		t.Errorf("stats[1].NumberCount = %d, want 3", stats[1].NumberCount)
	}

	// Check Active column
	if stats[2].InferredType != CellTypeBool {
		t.Errorf("stats[2].InferredType = %v, want CellTypeBool", stats[2].InferredType)
	}
	if stats[2].BoolCount != 3 {
		t.Errorf("stats[2].BoolCount = %d, want 3", stats[2].BoolCount)
	}
}

func TestTable_AnalyzeColumns_EmptyTable(t *testing.T) {
	table := Table{
		Headers: []string{},
		Rows:    []Row{},
	}

	stats := table.AnalyzeColumns()

	if stats != nil {
		t.Errorf("AnalyzeColumns() on empty table should return nil, got %v", stats)
	}
}

func TestTable_AnalyzeColumns_WithEmptyCells(t *testing.T) {
	table := Table{
		Headers: []string{"Name", "Value"},
		Rows: []Row{
			{
				Values: map[string]Cell{
					"Name":  {Value: "Item1", Type: CellTypeString, RawValue: "Item1"},
					"Value": {Value: 100.0, Type: CellTypeNumber, RawValue: "100"},
				},
			},
			{
				Values: map[string]Cell{
					"Name":  {Value: "", Type: CellTypeEmpty, RawValue: ""},
					"Value": {Value: nil, Type: CellTypeEmpty, RawValue: ""},
				},
			},
			{
				Values: map[string]Cell{
					"Name":  {Value: "Item2", Type: CellTypeString, RawValue: "Item2"},
					"Value": {Value: 200.0, Type: CellTypeNumber, RawValue: "200"},
				},
			},
		},
	}

	stats := table.AnalyzeColumns()

	// Name column should have 1 empty
	if stats[0].EmptyCount != 1 {
		t.Errorf("stats[0].EmptyCount = %d, want 1", stats[0].EmptyCount)
	}
	if stats[0].NullCount != 1 {
		t.Errorf("stats[0].NullCount = %d, want 1", stats[0].NullCount)
	}

	// Value column should have 1 empty
	if stats[1].EmptyCount != 1 {
		t.Errorf("stats[1].EmptyCount = %d, want 1", stats[1].EmptyCount)
	}
}

func TestTable_AnalyzeColumns_AllEmpty(t *testing.T) {
	table := Table{
		Headers: []string{"Empty"},
		Rows: []Row{
			{Values: map[string]Cell{"Empty": {Type: CellTypeEmpty, RawValue: ""}}},
			{Values: map[string]Cell{"Empty": {Type: CellTypeEmpty, RawValue: ""}}},
		},
	}

	stats := table.AnalyzeColumns()

	if stats[0].InferredType != CellTypeEmpty {
		t.Errorf("All-empty column should infer CellTypeEmpty, got %v", stats[0].InferredType)
	}
}

func TestTable_AnalyzeColumns_SampleValues(t *testing.T) {
	table := Table{
		Headers: []string{"ID"},
		Rows:    []Row{},
	}

	// Add 10 rows
	for i := 1; i <= 10; i++ {
		table.Rows = append(table.Rows, Row{
			Values: map[string]Cell{
				"ID": {Value: float64(i), Type: CellTypeNumber, RawValue: string(rune('0' + i))},
			},
		})
	}

	stats := table.AnalyzeColumns()

	// Should only have 5 sample values
	if len(stats[0].SampleValues) != 5 {
		t.Errorf("SampleValues should have max 5 items, got %d", len(stats[0].SampleValues))
	}
}

func TestTable_AnalyzeColumns_MixedTypes(t *testing.T) {
	table := Table{
		Headers: []string{"Mixed"},
		Rows: []Row{
			{Values: map[string]Cell{"Mixed": {Value: "text", Type: CellTypeString, RawValue: "text"}}},
			{Values: map[string]Cell{"Mixed": {Value: 123.0, Type: CellTypeNumber, RawValue: "123"}}},
			{Values: map[string]Cell{"Mixed": {Value: 456.0, Type: CellTypeNumber, RawValue: "456"}}},
			{Values: map[string]Cell{"Mixed": {Value: 789.0, Type: CellTypeNumber, RawValue: "789"}}},
		},
	}

	stats := table.AnalyzeColumns()

	// Numbers are more common, so should infer Number
	if stats[0].InferredType != CellTypeNumber {
		t.Errorf("Mixed column with more numbers should infer CellTypeNumber, got %v", stats[0].InferredType)
	}
	if stats[0].StringCount != 1 {
		t.Errorf("stats[0].StringCount = %d, want 1", stats[0].StringCount)
	}
	if stats[0].NumberCount != 3 {
		t.Errorf("stats[0].NumberCount = %d, want 3", stats[0].NumberCount)
	}
}

func TestTable_AnalyzeColumns_DateType(t *testing.T) {
	table := Table{
		Headers: []string{"Date"},
		Rows: []Row{
			{Values: map[string]Cell{"Date": {Type: CellTypeDate, RawValue: "2025-01-01"}}},
			{Values: map[string]Cell{"Date": {Type: CellTypeDate, RawValue: "2025-01-02"}}},
			{Values: map[string]Cell{"Date": {Type: CellTypeDate, RawValue: "2025-01-03"}}},
		},
	}

	stats := table.AnalyzeColumns()

	if stats[0].InferredType != CellTypeDate {
		t.Errorf("Date column should infer CellTypeDate, got %v", stats[0].InferredType)
	}
	if stats[0].DateCount != 3 {
		t.Errorf("stats[0].DateCount = %d, want 3", stats[0].DateCount)
	}
}

func TestTable_AnalyzeColumns_MissingCell(t *testing.T) {
	table := Table{
		Headers: []string{"Col1", "Col2"},
		Rows: []Row{
			{Values: map[string]Cell{"Col1": {Value: "A", Type: CellTypeString, RawValue: "A"}}},
			// Col2 is missing from this row's Values map
		},
	}

	stats := table.AnalyzeColumns()

	// Col2 should show as empty since it's missing from Values
	if stats[1].EmptyCount != 1 {
		t.Errorf("Missing cell should count as empty, got EmptyCount = %d", stats[1].EmptyCount)
	}
}

func TestTable_AnalyzeColumns_NumericStats(t *testing.T) {
	table := Table{
		Headers: []string{"Amount"},
		Rows: []Row{
			{Values: map[string]Cell{"Amount": {Value: 10.0, Type: CellTypeNumber, RawValue: "10"}}},
			{Values: map[string]Cell{"Amount": {Value: 20.0, Type: CellTypeNumber, RawValue: "20"}}},
			{Values: map[string]Cell{"Amount": {Value: 30.0, Type: CellTypeNumber, RawValue: "30"}}},
			{Values: map[string]Cell{"Amount": {Value: 40.0, Type: CellTypeNumber, RawValue: "40"}}},
		},
	}

	stats := table.AnalyzeColumns()

	if !stats[0].HasNumericStats {
		t.Error("HasNumericStats should be true for numeric column")
	}
	if stats[0].Min != 10.0 {
		t.Errorf("Min = %v, want 10.0", stats[0].Min)
	}
	if stats[0].Max != 40.0 {
		t.Errorf("Max = %v, want 40.0", stats[0].Max)
	}
	if stats[0].Sum != 100.0 {
		t.Errorf("Sum = %v, want 100.0", stats[0].Sum)
	}
	if stats[0].Avg != 25.0 {
		t.Errorf("Avg = %v, want 25.0", stats[0].Avg)
	}
}

func TestTable_AnalyzeColumns_NumericStats_WithNegatives(t *testing.T) {
	table := Table{
		Headers: []string{"Value"},
		Rows: []Row{
			{Values: map[string]Cell{"Value": {Value: -50.0, Type: CellTypeNumber, RawValue: "-50"}}},
			{Values: map[string]Cell{"Value": {Value: 0.0, Type: CellTypeNumber, RawValue: "0"}}},
			{Values: map[string]Cell{"Value": {Value: 50.0, Type: CellTypeNumber, RawValue: "50"}}},
		},
	}

	stats := table.AnalyzeColumns()

	if !stats[0].HasNumericStats {
		t.Error("HasNumericStats should be true")
	}
	if stats[0].Min != -50.0 {
		t.Errorf("Min = %v, want -50.0", stats[0].Min)
	}
	if stats[0].Max != 50.0 {
		t.Errorf("Max = %v, want 50.0", stats[0].Max)
	}
	if stats[0].Sum != 0.0 {
		t.Errorf("Sum = %v, want 0.0", stats[0].Sum)
	}
	if stats[0].Avg != 0.0 {
		t.Errorf("Avg = %v, want 0.0", stats[0].Avg)
	}
}

func TestTable_AnalyzeColumns_NumericStats_SingleValue(t *testing.T) {
	table := Table{
		Headers: []string{"Single"},
		Rows: []Row{
			{Values: map[string]Cell{"Single": {Value: 42.5, Type: CellTypeNumber, RawValue: "42.5"}}},
		},
	}

	stats := table.AnalyzeColumns()

	if !stats[0].HasNumericStats {
		t.Error("HasNumericStats should be true")
	}
	if stats[0].Min != 42.5 {
		t.Errorf("Min = %v, want 42.5", stats[0].Min)
	}
	if stats[0].Max != 42.5 {
		t.Errorf("Max = %v, want 42.5", stats[0].Max)
	}
	if stats[0].Sum != 42.5 {
		t.Errorf("Sum = %v, want 42.5", stats[0].Sum)
	}
	if stats[0].Avg != 42.5 {
		t.Errorf("Avg = %v, want 42.5", stats[0].Avg)
	}
}

func TestTable_AnalyzeColumns_NumericStats_NoNumericValues(t *testing.T) {
	table := Table{
		Headers: []string{"Text"},
		Rows: []Row{
			{Values: map[string]Cell{"Text": {Value: "hello", Type: CellTypeString, RawValue: "hello"}}},
			{Values: map[string]Cell{"Text": {Value: "world", Type: CellTypeString, RawValue: "world"}}},
		},
	}

	stats := table.AnalyzeColumns()

	if stats[0].HasNumericStats {
		t.Error("HasNumericStats should be false for non-numeric column")
	}
	if stats[0].Min != 0 {
		t.Errorf("Min should be 0 for non-numeric column, got %v", stats[0].Min)
	}
	if stats[0].Max != 0 {
		t.Errorf("Max should be 0 for non-numeric column, got %v", stats[0].Max)
	}
	if stats[0].Sum != 0 {
		t.Errorf("Sum should be 0 for non-numeric column, got %v", stats[0].Sum)
	}
	if stats[0].Avg != 0 {
		t.Errorf("Avg should be 0 for non-numeric column, got %v", stats[0].Avg)
	}
}

func TestTable_AnalyzeColumns_NumericStats_MixedColumn(t *testing.T) {
	table := Table{
		Headers: []string{"Mixed"},
		Rows: []Row{
			{Values: map[string]Cell{"Mixed": {Value: "text", Type: CellTypeString, RawValue: "text"}}},
			{Values: map[string]Cell{"Mixed": {Value: 10.0, Type: CellTypeNumber, RawValue: "10"}}},
			{Values: map[string]Cell{"Mixed": {Value: 20.0, Type: CellTypeNumber, RawValue: "20"}}},
			{Values: map[string]Cell{"Mixed": {Value: 30.0, Type: CellTypeNumber, RawValue: "30"}}},
		},
	}

	stats := table.AnalyzeColumns()

	// Should have numeric stats for the 3 numeric values
	if !stats[0].HasNumericStats {
		t.Error("HasNumericStats should be true (has some numeric values)")
	}
	if stats[0].Min != 10.0 {
		t.Errorf("Min = %v, want 10.0", stats[0].Min)
	}
	if stats[0].Max != 30.0 {
		t.Errorf("Max = %v, want 30.0", stats[0].Max)
	}
	if stats[0].Sum != 60.0 {
		t.Errorf("Sum = %v, want 60.0", stats[0].Sum)
	}
	if stats[0].Avg != 20.0 {
		t.Errorf("Avg = %v, want 20.0 (60/3)", stats[0].Avg)
	}
}

func TestTable_AnalyzeColumns_NumericStats_WithEmptyCells(t *testing.T) {
	table := Table{
		Headers: []string{"Values"},
		Rows: []Row{
			{Values: map[string]Cell{"Values": {Value: 100.0, Type: CellTypeNumber, RawValue: "100"}}},
			{Values: map[string]Cell{"Values": {Value: nil, Type: CellTypeEmpty, RawValue: ""}}},
			{Values: map[string]Cell{"Values": {Value: 200.0, Type: CellTypeNumber, RawValue: "200"}}},
			{Values: map[string]Cell{"Values": {Value: nil, Type: CellTypeEmpty, RawValue: ""}}},
			{Values: map[string]Cell{"Values": {Value: 300.0, Type: CellTypeNumber, RawValue: "300"}}},
		},
	}

	stats := table.AnalyzeColumns()

	if !stats[0].HasNumericStats {
		t.Error("HasNumericStats should be true")
	}
	// Stats should only consider the 3 numeric values
	if stats[0].NumberCount != 3 {
		t.Errorf("NumberCount = %v, want 3", stats[0].NumberCount)
	}
	if stats[0].Min != 100.0 {
		t.Errorf("Min = %v, want 100.0", stats[0].Min)
	}
	if stats[0].Max != 300.0 {
		t.Errorf("Max = %v, want 300.0", stats[0].Max)
	}
	if stats[0].Sum != 600.0 {
		t.Errorf("Sum = %v, want 600.0", stats[0].Sum)
	}
	if stats[0].Avg != 200.0 {
		t.Errorf("Avg = %v, want 200.0", stats[0].Avg)
	}
}

func TestTable_AnalyzeColumns_NumericStats_Decimals(t *testing.T) {
	table := Table{
		Headers: []string{"Price"},
		Rows: []Row{
			{Values: map[string]Cell{"Price": {Value: 9.99, Type: CellTypeNumber, RawValue: "9.99"}}},
			{Values: map[string]Cell{"Price": {Value: 19.99, Type: CellTypeNumber, RawValue: "19.99"}}},
			{Values: map[string]Cell{"Price": {Value: 29.99, Type: CellTypeNumber, RawValue: "29.99"}}},
		},
	}

	stats := table.AnalyzeColumns()

	if !stats[0].HasNumericStats {
		t.Error("HasNumericStats should be true")
	}
	if stats[0].Min != 9.99 {
		t.Errorf("Min = %v, want 9.99", stats[0].Min)
	}
	if stats[0].Max != 29.99 {
		t.Errorf("Max = %v, want 29.99", stats[0].Max)
	}
	expectedSum := 9.99 + 19.99 + 29.99
	if stats[0].Sum != expectedSum {
		t.Errorf("Sum = %v, want %v", stats[0].Sum, expectedSum)
	}
	expectedAvg := expectedSum / 3
	if stats[0].Avg != expectedAvg {
		t.Errorf("Avg = %v, want %v", stats[0].Avg, expectedAvg)
	}
}

// =============================================================================
// DetectionConfig Tests
// =============================================================================

func TestDefaultConfig(t *testing.T) {
	config := DefaultConfig()

	if config.MinColumns != 2 {
		t.Errorf("DefaultConfig().MinColumns = %v, want 2", config.MinColumns)
	}
	if config.MinRows != 2 {
		t.Errorf("DefaultConfig().MinRows = %v, want 2", config.MinRows)
	}
	if config.MaxEmptyRows != 2 {
		t.Errorf("DefaultConfig().MaxEmptyRows = %v, want 2", config.MaxEmptyRows)
	}
	if config.HeaderDensity != 0.5 {
		t.Errorf("DefaultConfig().HeaderDensity = %v, want 0.5", config.HeaderDensity)
	}
	if config.ColumnConsistency != 0.7 {
		t.Errorf("DefaultConfig().ColumnConsistency = %v, want 0.7", config.ColumnConsistency)
	}
}

// =============================================================================
// CellType Tests
// =============================================================================

func TestCellType_Values(t *testing.T) {
	// Ensure cell types have expected values
	if CellTypeEmpty != 0 {
		t.Errorf("CellTypeEmpty = %v, want 0", CellTypeEmpty)
	}
	if CellTypeString != 1 {
		t.Errorf("CellTypeString = %v, want 1", CellTypeString)
	}
	if CellTypeNumber != 2 {
		t.Errorf("CellTypeNumber = %v, want 2", CellTypeNumber)
	}
	if CellTypeDate != 3 {
		t.Errorf("CellTypeDate = %v, want 3", CellTypeDate)
	}
	if CellTypeBool != 4 {
		t.Errorf("CellTypeBool = %v, want 4", CellTypeBool)
	}
	if CellTypeFormula != 5 {
		t.Errorf("CellTypeFormula = %v, want 5", CellTypeFormula)
	}
}

// =============================================================================
// TableBoundary Tests
// =============================================================================

func TestTableBoundary_Dimensions(t *testing.T) {
	tests := []struct {
		name         string
		boundary     TableBoundary
		expectedRows int
		expectedCols int
	}{
		{
			name:         "single cell",
			boundary:     TableBoundary{StartRow: 0, EndRow: 0, StartCol: 0, EndCol: 0},
			expectedRows: 1,
			expectedCols: 1,
		},
		{
			name:         "3x4 table",
			boundary:     TableBoundary{StartRow: 0, EndRow: 2, StartCol: 0, EndCol: 3},
			expectedRows: 3,
			expectedCols: 4,
		},
		{
			name:         "offset table",
			boundary:     TableBoundary{StartRow: 5, EndRow: 10, StartCol: 2, EndCol: 7},
			expectedRows: 6,
			expectedCols: 6,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			rows := tt.boundary.EndRow - tt.boundary.StartRow + 1
			cols := tt.boundary.EndCol - tt.boundary.StartCol + 1
			if rows != tt.expectedRows {
				t.Errorf("row count = %v, want %v", rows, tt.expectedRows)
			}
			if cols != tt.expectedCols {
				t.Errorf("col count = %v, want %v", cols, tt.expectedCols)
			}
		})
	}
}

// =============================================================================
// Workbook and Sheet Tests
// =============================================================================

func TestWorkbook_Structure(t *testing.T) {
	wb := Workbook{
		FilePath: "/path/to/file.xlsx",
		Sheets: []Sheet{
			{
				Name:  "Sheet1",
				Index: 0,
				Tables: []Table{
					{Name: "Table1", Headers: []string{"A", "B"}},
				},
			},
			{
				Name:  "Sheet2",
				Index: 1,
				Tables: []Table{
					{Name: "Table2", Headers: []string{"C", "D", "E"}},
					{Name: "Table3", Headers: []string{"F"}},
				},
			},
		},
	}

	if wb.FilePath != "/path/to/file.xlsx" {
		t.Errorf("FilePath = %v, want /path/to/file.xlsx", wb.FilePath)
	}

	if len(wb.Sheets) != 2 {
		t.Errorf("len(Sheets) = %v, want 2", len(wb.Sheets))
	}

	if len(wb.Sheets[0].Tables) != 1 {
		t.Errorf("len(Sheets[0].Tables) = %v, want 1", len(wb.Sheets[0].Tables))
	}

	if len(wb.Sheets[1].Tables) != 2 {
		t.Errorf("len(Sheets[1].Tables) = %v, want 2", len(wb.Sheets[1].Tables))
	}

	// Count total tables
	totalTables := 0
	for _, sheet := range wb.Sheets {
		totalTables += len(sheet.Tables)
	}
	if totalTables != 3 {
		t.Errorf("total tables = %v, want 3", totalTables)
	}
}

// =============================================================================
// DiffTables Tests
// =============================================================================

func TestDiffTables_AddedRows(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
		},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if len(result.AddedRows) != 1 {
		t.Errorf("Expected 1 added row, got %d", len(result.AddedRows))
	}
	if len(result.RemovedRows) != 0 {
		t.Errorf("Expected 0 removed rows, got %d", len(result.RemovedRows))
	}
	if len(result.ModifiedRows) != 0 {
		t.Errorf("Expected 0 modified rows, got %d", len(result.ModifiedRows))
	}

	// Verify added row
	if cell, ok := result.AddedRows[0].Get("ID"); ok {
		if cell.RawValue != "3" {
			t.Errorf("Expected added row ID '3', got '%s'", cell.RawValue)
		}
	}
}

func TestDiffTables_RemovedRows(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}}},
		},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if len(result.AddedRows) != 0 {
		t.Errorf("Expected 0 added rows, got %d", len(result.AddedRows))
	}
	if len(result.RemovedRows) != 1 {
		t.Errorf("Expected 1 removed row, got %d", len(result.RemovedRows))
	}

	// Verify removed row
	if cell, ok := result.RemovedRows[0].Get("ID"); ok {
		if cell.RawValue != "2" {
			t.Errorf("Expected removed row ID '2', got '%s'", cell.RawValue)
		}
	}
}

func TestDiffTables_ModifiedRows(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name", "Email"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}, "Email": {RawValue: "alice@old.com"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}, "Email": {RawValue: "bob@test.com"}}},
		},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name", "Email"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}, "Email": {RawValue: "alice@new.com"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}, "Email": {RawValue: "bob@test.com"}}},
		},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if len(result.AddedRows) != 0 {
		t.Errorf("Expected 0 added rows, got %d", len(result.AddedRows))
	}
	if len(result.RemovedRows) != 0 {
		t.Errorf("Expected 0 removed rows, got %d", len(result.RemovedRows))
	}
	if len(result.ModifiedRows) != 1 {
		t.Errorf("Expected 1 modified row, got %d", len(result.ModifiedRows))
	}

	// Verify modified row
	if result.ModifiedRows[0].KeyValue != "1" {
		t.Errorf("Expected modified row key '1', got '%s'", result.ModifiedRows[0].KeyValue)
	}
	if len(result.ModifiedRows[0].Changes) != 1 {
		t.Errorf("Expected 1 change, got %d", len(result.ModifiedRows[0].Changes))
	}
	if result.ModifiedRows[0].Changes[0].Column != "Email" {
		t.Errorf("Expected change in 'Email' column, got '%s'", result.ModifiedRows[0].Changes[0].Column)
	}
	if result.ModifiedRows[0].Changes[0].OldValue != "alice@old.com" {
		t.Errorf("Expected old value 'alice@old.com', got '%s'", result.ModifiedRows[0].Changes[0].OldValue)
	}
	if result.ModifiedRows[0].Changes[0].NewValue != "alice@new.com" {
		t.Errorf("Expected new value 'alice@new.com', got '%s'", result.ModifiedRows[0].Changes[0].NewValue)
	}
}

func TestDiffTables_MultipleChanges(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name", "Age", "Status"},
		Rows: []Row{
			{Values: map[string]Cell{
				"ID":     {RawValue: "1"},
				"Name":   {RawValue: "Alice"},
				"Age":    {RawValue: "25"},
				"Status": {RawValue: "active"},
			}},
		},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name", "Age", "Status"},
		Rows: []Row{
			{Values: map[string]Cell{
				"ID":     {RawValue: "1"},
				"Name":   {RawValue: "Alice Smith"},
				"Age":    {RawValue: "26"},
				"Status": {RawValue: "active"},
			}},
		},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if len(result.ModifiedRows) != 1 {
		t.Fatalf("Expected 1 modified row, got %d", len(result.ModifiedRows))
	}
	if len(result.ModifiedRows[0].Changes) != 2 {
		t.Errorf("Expected 2 changes (Name and Age), got %d", len(result.ModifiedRows[0].Changes))
	}
}

func TestDiffTables_NoChanges(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
		},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
		},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if result.HasChanges() {
		t.Errorf("Expected no changes, but HasChanges() returned true")
	}
	if result.TotalChanges() != 0 {
		t.Errorf("Expected TotalChanges() = 0, got %d", result.TotalChanges())
	}
}

func TestDiffTables_EmptyTables(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows:    []Row{},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows:    []Row{},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if result.HasChanges() {
		t.Errorf("Expected no changes for empty tables")
	}
}

func TestDiffTables_OldEmpty(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows:    []Row{},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
		},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if len(result.AddedRows) != 2 {
		t.Errorf("Expected 2 added rows, got %d", len(result.AddedRows))
	}
	if len(result.RemovedRows) != 0 {
		t.Errorf("Expected 0 removed rows, got %d", len(result.RemovedRows))
	}
}

func TestDiffTables_NewEmpty(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}}},
		},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name"},
		Rows:    []Row{},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if len(result.AddedRows) != 0 {
		t.Errorf("Expected 0 added rows, got %d", len(result.AddedRows))
	}
	if len(result.RemovedRows) != 2 {
		t.Errorf("Expected 2 removed rows, got %d", len(result.RemovedRows))
	}
}

func TestDiffTables_AllChangeTypes(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"ID", "Name", "Status"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}, "Status": {RawValue: "active"}}},
			{Values: map[string]Cell{"ID": {RawValue: "2"}, "Name": {RawValue: "Bob"}, "Status": {RawValue: "active"}}},
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}, "Status": {RawValue: "active"}}},
		},
	}

	newTable := &Table{
		Headers: []string{"ID", "Name", "Status"},
		Rows: []Row{
			{Values: map[string]Cell{"ID": {RawValue: "1"}, "Name": {RawValue: "Alice"}, "Status": {RawValue: "inactive"}}}, // Modified
			// ID 2 removed
			{Values: map[string]Cell{"ID": {RawValue: "3"}, "Name": {RawValue: "Charlie"}, "Status": {RawValue: "active"}}}, // Unchanged
			{Values: map[string]Cell{"ID": {RawValue: "4"}, "Name": {RawValue: "Diana"}, "Status": {RawValue: "active"}}},   // Added
		},
	}

	result := DiffTables(oldTable, newTable, "ID")

	if len(result.AddedRows) != 1 {
		t.Errorf("Expected 1 added row, got %d", len(result.AddedRows))
	}
	if len(result.RemovedRows) != 1 {
		t.Errorf("Expected 1 removed row, got %d", len(result.RemovedRows))
	}
	if len(result.ModifiedRows) != 1 {
		t.Errorf("Expected 1 modified row, got %d", len(result.ModifiedRows))
	}
	if result.TotalChanges() != 3 {
		t.Errorf("Expected TotalChanges() = 3, got %d", result.TotalChanges())
	}
}

func TestDiffResult_HasChanges(t *testing.T) {
	tests := []struct {
		name     string
		result   DiffResult
		expected bool
	}{
		{
			name:     "no changes",
			result:   DiffResult{AddedRows: []Row{}, RemovedRows: []Row{}, ModifiedRows: []RowDiff{}},
			expected: false,
		},
		{
			name:     "has added",
			result:   DiffResult{AddedRows: []Row{{}}, RemovedRows: []Row{}, ModifiedRows: []RowDiff{}},
			expected: true,
		},
		{
			name:     "has removed",
			result:   DiffResult{AddedRows: []Row{}, RemovedRows: []Row{{}}, ModifiedRows: []RowDiff{}},
			expected: true,
		},
		{
			name:     "has modified",
			result:   DiffResult{AddedRows: []Row{}, RemovedRows: []Row{}, ModifiedRows: []RowDiff{{}}},
			expected: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.result.HasChanges(); got != tt.expected {
				t.Errorf("HasChanges() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestDiffResult_TotalChanges(t *testing.T) {
	result := DiffResult{
		AddedRows:    []Row{{}, {}},
		RemovedRows:  []Row{{}},
		ModifiedRows: []RowDiff{{}, {}, {}},
	}

	if result.TotalChanges() != 6 {
		t.Errorf("TotalChanges() = %d, want 6", result.TotalChanges())
	}
}

func TestDiffTables_KeyColumn(t *testing.T) {
	oldTable := &Table{
		Headers: []string{"Email", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice"}}},
		},
	}

	newTable := &Table{
		Headers: []string{"Email", "Name"},
		Rows: []Row{
			{Values: map[string]Cell{"Email": {RawValue: "alice@test.com"}, "Name": {RawValue: "Alice Updated"}}},
		},
	}

	result := DiffTables(oldTable, newTable, "Email")

	if result.KeyColumn != "Email" {
		t.Errorf("Expected KeyColumn 'Email', got '%s'", result.KeyColumn)
	}
	if len(result.ModifiedRows) != 1 {
		t.Errorf("Expected 1 modified row using Email as key")
	}
}
