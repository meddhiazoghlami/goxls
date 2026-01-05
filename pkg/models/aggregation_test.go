package models

import (
	"testing"
)

// Helper function to create a test table
func createTestTable() *Table {
	return &Table{
		Name:    "TestTable",
		Headers: []string{"Category", "Product", "Price", "Quantity"},
		Rows: []Row{
			{
				Index: 0,
				Values: map[string]Cell{
					"Category": {Value: "Electronics", Type: CellTypeString, RawValue: "Electronics"},
					"Product":  {Value: "Phone", Type: CellTypeString, RawValue: "Phone"},
					"Price":    {Value: 999.99, Type: CellTypeNumber, RawValue: "999.99"},
					"Quantity": {Value: 10.0, Type: CellTypeNumber, RawValue: "10"},
				},
				Cells: []Cell{
					{Value: "Electronics", Type: CellTypeString, RawValue: "Electronics"},
					{Value: "Phone", Type: CellTypeString, RawValue: "Phone"},
					{Value: 999.99, Type: CellTypeNumber, RawValue: "999.99"},
					{Value: 10.0, Type: CellTypeNumber, RawValue: "10"},
				},
			},
			{
				Index: 1,
				Values: map[string]Cell{
					"Category": {Value: "Electronics", Type: CellTypeString, RawValue: "Electronics"},
					"Product":  {Value: "Laptop", Type: CellTypeString, RawValue: "Laptop"},
					"Price":    {Value: 1499.99, Type: CellTypeNumber, RawValue: "1499.99"},
					"Quantity": {Value: 5.0, Type: CellTypeNumber, RawValue: "5"},
				},
				Cells: []Cell{
					{Value: "Electronics", Type: CellTypeString, RawValue: "Electronics"},
					{Value: "Laptop", Type: CellTypeString, RawValue: "Laptop"},
					{Value: 1499.99, Type: CellTypeNumber, RawValue: "1499.99"},
					{Value: 5.0, Type: CellTypeNumber, RawValue: "5"},
				},
			},
			{
				Index: 2,
				Values: map[string]Cell{
					"Category": {Value: "Clothing", Type: CellTypeString, RawValue: "Clothing"},
					"Product":  {Value: "Shirt", Type: CellTypeString, RawValue: "Shirt"},
					"Price":    {Value: 29.99, Type: CellTypeNumber, RawValue: "29.99"},
					"Quantity": {Value: 100.0, Type: CellTypeNumber, RawValue: "100"},
				},
				Cells: []Cell{
					{Value: "Clothing", Type: CellTypeString, RawValue: "Clothing"},
					{Value: "Shirt", Type: CellTypeString, RawValue: "Shirt"},
					{Value: 29.99, Type: CellTypeNumber, RawValue: "29.99"},
					{Value: 100.0, Type: CellTypeNumber, RawValue: "100"},
				},
			},
			{
				Index: 3,
				Values: map[string]Cell{
					"Category": {Value: "Clothing", Type: CellTypeString, RawValue: "Clothing"},
					"Product":  {Value: "Pants", Type: CellTypeString, RawValue: "Pants"},
					"Price":    {Value: 49.99, Type: CellTypeNumber, RawValue: "49.99"},
					"Quantity": {Value: 50.0, Type: CellTypeNumber, RawValue: "50"},
				},
				Cells: []Cell{
					{Value: "Clothing", Type: CellTypeString, RawValue: "Clothing"},
					{Value: "Pants", Type: CellTypeString, RawValue: "Pants"},
					{Value: 49.99, Type: CellTypeNumber, RawValue: "49.99"},
					{Value: 50.0, Type: CellTypeNumber, RawValue: "50"},
				},
			},
		},
	}
}

func TestAggregateOp_String(t *testing.T) {
	tests := []struct {
		op       AggregateOp
		expected string
	}{
		{AggSum, "Sum"},
		{AggCount, "Count"},
		{AggAvg, "Avg"},
		{AggMin, "Min"},
		{AggMax, "Max"},
		{AggregateOp(99), "Unknown"},
	}

	for _, tt := range tests {
		t.Run(tt.expected, func(t *testing.T) {
			if got := tt.op.String(); got != tt.expected {
				t.Errorf("AggregateOp.String() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestAggregateFunc_OutputName(t *testing.T) {
	tests := []struct {
		name     string
		agg      AggregateFunc
		expected string
	}{
		{"Sum without alias", Sum("Amount"), "Sum_Amount"},
		{"Count without alias", Count("ID"), "Count_ID"},
		{"Avg without alias", Avg("Price"), "Avg_Price"},
		{"Min without alias", Min("Value"), "Min_Value"},
		{"Max without alias", Max("Score"), "Max_Score"},
		{"Sum with alias", Sum("Amount").As("TotalAmount"), "TotalAmount"},
		{"Count with alias", Count("ID").As("NumRecords"), "NumRecords"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.agg.OutputName(); got != tt.expected {
				t.Errorf("AggregateFunc.OutputName() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestGroupBy_Sum(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate(Sum("Price"))

	if len(result.Headers) != 2 {
		t.Errorf("Expected 2 headers, got %d", len(result.Headers))
	}

	if result.Headers[0] != "Category" {
		t.Errorf("Expected first header to be 'Category', got %s", result.Headers[0])
	}

	if result.Headers[1] != "Sum_Price" {
		t.Errorf("Expected second header to be 'Sum_Price', got %s", result.Headers[1])
	}

	if len(result.Rows) != 2 {
		t.Errorf("Expected 2 rows (2 categories), got %d", len(result.Rows))
	}

	// Check sums - Electronics: 999.99 + 1499.99 = 2499.98, Clothing: 29.99 + 49.99 = 79.98
	for _, row := range result.Rows {
		category, _ := row.Get("Category")
		sumPrice, _ := row.Get("Sum_Price")
		val, _ := sumPrice.AsFloat()

		switch category.RawValue {
		case "Electronics":
			expected := 2499.98
			if val != expected {
				t.Errorf("Electronics sum = %v, want %v", val, expected)
			}
		case "Clothing":
			expected := 79.98
			if val != expected {
				t.Errorf("Clothing sum = %v, want %v", val, expected)
			}
		}
	}
}

func TestGroupBy_Count(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate(Count("Product"))

	if len(result.Rows) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(result.Rows))
	}

	for _, row := range result.Rows {
		countCell, _ := row.Get("Count_Product")
		count, _ := countCell.AsFloat()

		if count != 2 {
			t.Errorf("Expected count of 2 for each category, got %v", count)
		}
	}
}

func TestGroupBy_Avg(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate(Avg("Price"))

	for _, row := range result.Rows {
		category, _ := row.Get("Category")
		avgCell, _ := row.Get("Avg_Price")
		avg, _ := avgCell.AsFloat()

		switch category.RawValue {
		case "Electronics":
			expected := (999.99 + 1499.99) / 2
			if avg != expected {
				t.Errorf("Electronics avg = %v, want %v", avg, expected)
			}
		case "Clothing":
			expected := (29.99 + 49.99) / 2
			if avg != expected {
				t.Errorf("Clothing avg = %v, want %v", avg, expected)
			}
		}
	}
}

func TestGroupBy_Min(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate(Min("Price"))

	for _, row := range result.Rows {
		category, _ := row.Get("Category")
		minCell, _ := row.Get("Min_Price")
		min, _ := minCell.AsFloat()

		switch category.RawValue {
		case "Electronics":
			expected := 999.99
			if min != expected {
				t.Errorf("Electronics min = %v, want %v", min, expected)
			}
		case "Clothing":
			expected := 29.99
			if min != expected {
				t.Errorf("Clothing min = %v, want %v", min, expected)
			}
		}
	}
}

func TestGroupBy_Max(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate(Max("Price"))

	for _, row := range result.Rows {
		category, _ := row.Get("Category")
		maxCell, _ := row.Get("Max_Price")
		max, _ := maxCell.AsFloat()

		switch category.RawValue {
		case "Electronics":
			expected := 1499.99
			if max != expected {
				t.Errorf("Electronics max = %v, want %v", max, expected)
			}
		case "Clothing":
			expected := 49.99
			if max != expected {
				t.Errorf("Clothing max = %v, want %v", max, expected)
			}
		}
	}
}

func TestGroupBy_MultipleAggregations(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate(
		Sum("Price").As("TotalPrice"),
		Count("Product").As("NumProducts"),
		Avg("Quantity"),
		Min("Price"),
		Max("Price"),
	)

	// Check headers
	expectedHeaders := []string{"Category", "TotalPrice", "NumProducts", "Avg_Quantity", "Min_Price", "Max_Price"}
	if len(result.Headers) != len(expectedHeaders) {
		t.Errorf("Expected %d headers, got %d", len(expectedHeaders), len(result.Headers))
	}

	for i, h := range expectedHeaders {
		if result.Headers[i] != h {
			t.Errorf("Header[%d] = %s, want %s", i, result.Headers[i], h)
		}
	}

	// Check we have 2 groups
	if len(result.Rows) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(result.Rows))
	}
}

func TestGroupBy_EmptyTable(t *testing.T) {
	table := &Table{
		Name:    "Empty",
		Headers: []string{"Category", "Price"},
		Rows:    []Row{},
	}

	result := table.GroupBy("Category").Aggregate(Sum("Price"))

	if len(result.Rows) != 0 {
		t.Errorf("Expected 0 rows for empty table, got %d", len(result.Rows))
	}

	if len(result.Headers) != 2 {
		t.Errorf("Expected 2 headers, got %d", len(result.Headers))
	}
}

func TestGroupBy_NoFuncs(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate()

	if len(result.Rows) != 0 {
		t.Errorf("Expected 0 rows when no aggregation funcs, got %d", len(result.Rows))
	}
}

func TestGroupBy_NonExistentColumn(t *testing.T) {
	table := createTestTable()

	// Aggregating a non-existent column should produce empty/zero results
	result := table.GroupBy("Category").Aggregate(Sum("NonExistent"))

	if len(result.Rows) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(result.Rows))
	}

	// The sum should be empty (no values found)
	for _, row := range result.Rows {
		cell, _ := row.Get("Sum_NonExistent")
		if cell.Type != CellTypeEmpty {
			t.Errorf("Expected empty cell for non-existent column, got type %v", cell.Type)
		}
	}
}

func TestGroupBy_MultipleGroupColumns(t *testing.T) {
	// Create a table with more granular data
	table := &Table{
		Name:    "Sales",
		Headers: []string{"Region", "Category", "Amount"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Region":   {Value: "North", Type: CellTypeString, RawValue: "North"},
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Amount":   {Value: 100.0, Type: CellTypeNumber, RawValue: "100"},
			}},
			{Index: 1, Values: map[string]Cell{
				"Region":   {Value: "North", Type: CellTypeString, RawValue: "North"},
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Amount":   {Value: 150.0, Type: CellTypeNumber, RawValue: "150"},
			}},
			{Index: 2, Values: map[string]Cell{
				"Region":   {Value: "North", Type: CellTypeString, RawValue: "North"},
				"Category": {Value: "B", Type: CellTypeString, RawValue: "B"},
				"Amount":   {Value: 200.0, Type: CellTypeNumber, RawValue: "200"},
			}},
			{Index: 3, Values: map[string]Cell{
				"Region":   {Value: "South", Type: CellTypeString, RawValue: "South"},
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Amount":   {Value: 300.0, Type: CellTypeNumber, RawValue: "300"},
			}},
		},
	}

	result := table.GroupBy("Region", "Category").Aggregate(Sum("Amount"))

	// Should have 3 groups: North/A, North/B, South/A
	if len(result.Rows) != 3 {
		t.Errorf("Expected 3 groups, got %d", len(result.Rows))
	}

	// Check headers
	if result.Headers[0] != "Region" || result.Headers[1] != "Category" || result.Headers[2] != "Sum_Amount" {
		t.Errorf("Unexpected headers: %v", result.Headers)
	}
}

func TestGroupBy_PreservesTableName(t *testing.T) {
	table := createTestTable()
	table.Name = "SalesData"

	result := table.GroupBy("Category").Aggregate(Sum("Price"))

	if result.Name != "SalesData" {
		t.Errorf("Expected table name 'SalesData', got '%s'", result.Name)
	}
}

func TestGroupBy_CountWithEmptyCells(t *testing.T) {
	table := &Table{
		Name:    "WithEmpty",
		Headers: []string{"Category", "Value"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: "X", Type: CellTypeString, RawValue: "X"},
			}},
			{Index: 1, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Type: CellTypeEmpty, RawValue: ""},
			}},
			{Index: 2, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: "Y", Type: CellTypeString, RawValue: "Y"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(Count("Value"))

	row := result.Rows[0]
	countCell, _ := row.Get("Count_Value")
	count, _ := countCell.AsFloat()

	// Should count only non-empty cells (2 out of 3)
	if count != 2 {
		t.Errorf("Expected count of 2 (excluding empty), got %v", count)
	}
}

func TestFormatFloat(t *testing.T) {
	tests := []struct {
		input    float64
		expected string
	}{
		{100.0, "100"},
		{100.5, "100.5"},
		{100.50, "100.5"},
		{0.0, "0"},
		{0.1, "0.1"},
		{0.10, "0.1"},
		{123.456, "123.456"},
	}

	for _, tt := range tests {
		t.Run(tt.expected, func(t *testing.T) {
			if got := formatFloat(tt.input); got != tt.expected {
				t.Errorf("formatFloat(%v) = %v, want %v", tt.input, got, tt.expected)
			}
		})
	}
}

func TestGroupBy_NilSource(t *testing.T) {
	// Test with nil source table
	grouped := &GroupedTable{
		source:    nil,
		groupCols: []string{"Category"},
	}

	result := grouped.Aggregate(Sum("Amount"))

	if result == nil {
		t.Fatal("Expected non-nil result")
	}
	if len(result.Rows) != 0 {
		t.Errorf("Expected 0 rows for nil source, got %d", len(result.Rows))
	}
	if len(result.Headers) != 0 {
		t.Errorf("Expected 0 headers for nil source, got %d", len(result.Headers))
	}
}

func TestGroupBy_NonExistentGroupColumn(t *testing.T) {
	table := createTestTable()

	// Group by a column that doesn't exist
	result := table.GroupBy("NonExistentColumn").Aggregate(Sum("Price"))

	// Should create one group with empty key
	if len(result.Rows) != 1 {
		t.Errorf("Expected 1 row (all rows grouped together), got %d", len(result.Rows))
	}

	// The group column should be empty
	row := result.Rows[0]
	groupCell, ok := row.Get("NonExistentColumn")
	if !ok {
		t.Error("Expected group column to exist in result")
	}
	if groupCell.Type != CellTypeEmpty {
		t.Errorf("Expected empty cell for non-existent group column, got type %v", groupCell.Type)
	}
}

func TestGroupBy_AvgWithNonNumericColumn(t *testing.T) {
	table := &Table{
		Name:    "Test",
		Headers: []string{"Category", "Name"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Name":     {Value: "Alice", Type: CellTypeString, RawValue: "Alice"},
			}},
			{Index: 1, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Name":     {Value: "Bob", Type: CellTypeString, RawValue: "Bob"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(Avg("Name"))

	row := result.Rows[0]
	avgCell, _ := row.Get("Avg_Name")

	// Avg of non-numeric column should be empty
	if avgCell.Type != CellTypeEmpty {
		t.Errorf("Expected empty cell for Avg of non-numeric column, got type %v", avgCell.Type)
	}
}

func TestGroupBy_MinWithNonNumericColumn(t *testing.T) {
	table := &Table{
		Name:    "Test",
		Headers: []string{"Category", "Name"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Name":     {Value: "Alice", Type: CellTypeString, RawValue: "Alice"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(Min("Name"))

	row := result.Rows[0]
	minCell, _ := row.Get("Min_Name")

	// Min of non-numeric column should be empty
	if minCell.Type != CellTypeEmpty {
		t.Errorf("Expected empty cell for Min of non-numeric column, got type %v", minCell.Type)
	}
}

func TestGroupBy_MaxWithNonNumericColumn(t *testing.T) {
	table := &Table{
		Name:    "Test",
		Headers: []string{"Category", "Name"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Name":     {Value: "Alice", Type: CellTypeString, RawValue: "Alice"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(Max("Name"))

	row := result.Rows[0]
	maxCell, _ := row.Get("Max_Name")

	// Max of non-numeric column should be empty
	if maxCell.Type != CellTypeEmpty {
		t.Errorf("Expected empty cell for Max of non-numeric column, got type %v", maxCell.Type)
	}
}

func TestGroupBy_SumWithNonNumericColumn(t *testing.T) {
	table := &Table{
		Name:    "Test",
		Headers: []string{"Category", "Name"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Name":     {Value: "Alice", Type: CellTypeString, RawValue: "Alice"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(Sum("Name"))

	row := result.Rows[0]
	sumCell, _ := row.Get("Sum_Name")

	// Sum of non-numeric column should be empty
	if sumCell.Type != CellTypeEmpty {
		t.Errorf("Expected empty cell for Sum of non-numeric column, got type %v", sumCell.Type)
	}
}

func TestGroupBy_MixedNumericAndNonNumeric(t *testing.T) {
	table := &Table{
		Name:    "Test",
		Headers: []string{"Category", "Value"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: 100.0, Type: CellTypeNumber, RawValue: "100"},
			}},
			{Index: 1, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: "NotANumber", Type: CellTypeString, RawValue: "NotANumber"},
			}},
			{Index: 2, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: 200.0, Type: CellTypeNumber, RawValue: "200"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(
		Sum("Value"),
		Avg("Value"),
		Min("Value"),
		Max("Value"),
		Count("Value"),
	)

	row := result.Rows[0]

	// Sum should be 300 (100 + 200, skipping non-numeric)
	sumCell, _ := row.Get("Sum_Value")
	sumVal, _ := sumCell.AsFloat()
	if sumVal != 300 {
		t.Errorf("Expected Sum=300, got %v", sumVal)
	}

	// Avg should be 150 (300 / 2 numeric values)
	avgCell, _ := row.Get("Avg_Value")
	avgVal, _ := avgCell.AsFloat()
	if avgVal != 150 {
		t.Errorf("Expected Avg=150, got %v", avgVal)
	}

	// Min should be 100
	minCell, _ := row.Get("Min_Value")
	minVal, _ := minCell.AsFloat()
	if minVal != 100 {
		t.Errorf("Expected Min=100, got %v", minVal)
	}

	// Max should be 200
	maxCell, _ := row.Get("Max_Value")
	maxVal, _ := maxCell.AsFloat()
	if maxVal != 200 {
		t.Errorf("Expected Max=200, got %v", maxVal)
	}

	// Count should be 3 (counts non-empty, including non-numeric)
	countCell, _ := row.Get("Count_Value")
	countVal, _ := countCell.AsFloat()
	if countVal != 3 {
		t.Errorf("Expected Count=3, got %v", countVal)
	}
}

func TestGroupBy_RowWithMissingGroupColumn(t *testing.T) {
	// Create a table where some rows don't have the group column
	table := &Table{
		Name:    "Test",
		Headers: []string{"Category", "Amount"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Amount":   {Value: 100.0, Type: CellTypeNumber, RawValue: "100"},
			}},
			{Index: 1, Values: map[string]Cell{
				// Missing "Category" key
				"Amount": {Value: 200.0, Type: CellTypeNumber, RawValue: "200"},
			}},
			{Index: 2, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Amount":   {Value: 300.0, Type: CellTypeNumber, RawValue: "300"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(Sum("Amount"))

	// Should have 2 groups: "A" and "" (empty for missing category)
	if len(result.Rows) != 2 {
		t.Errorf("Expected 2 groups, got %d", len(result.Rows))
	}
}

func TestGroupBy_AllAggregationsOnSameColumn(t *testing.T) {
	table := &Table{
		Name:    "Test",
		Headers: []string{"Category", "Value"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: 10.0, Type: CellTypeNumber, RawValue: "10"},
			}},
			{Index: 1, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: 20.0, Type: CellTypeNumber, RawValue: "20"},
			}},
			{Index: 2, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: 30.0, Type: CellTypeNumber, RawValue: "30"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(
		Sum("Value"),
		Count("Value"),
		Avg("Value"),
		Min("Value"),
		Max("Value"),
	)

	if len(result.Rows) != 1 {
		t.Fatalf("Expected 1 row, got %d", len(result.Rows))
	}

	row := result.Rows[0]

	// Verify all aggregations
	tests := []struct {
		column   string
		expected float64
	}{
		{"Sum_Value", 60},
		{"Count_Value", 3},
		{"Avg_Value", 20},
		{"Min_Value", 10},
		{"Max_Value", 30},
	}

	for _, tt := range tests {
		cell, ok := row.Get(tt.column)
		if !ok {
			t.Errorf("Column %s not found", tt.column)
			continue
		}
		val, ok := cell.AsFloat()
		if !ok {
			t.Errorf("Column %s is not numeric", tt.column)
			continue
		}
		if val != tt.expected {
			t.Errorf("%s = %v, want %v", tt.column, val, tt.expected)
		}
	}
}

func TestGroupBy_CellsArrayPopulated(t *testing.T) {
	table := createTestTable()

	result := table.GroupBy("Category").Aggregate(Sum("Price"), Count("Product"))

	for _, row := range result.Rows {
		// Verify Cells array is populated correctly
		if len(row.Cells) != 3 { // Category + Sum_Price + Count_Product
			t.Errorf("Expected 3 cells, got %d", len(row.Cells))
		}

		// Verify Values map matches Cells array
		if len(row.Values) != 3 {
			t.Errorf("Expected 3 values, got %d", len(row.Values))
		}
	}
}

func TestGroupBy_SingleRow(t *testing.T) {
	table := &Table{
		Name:    "Single",
		Headers: []string{"Category", "Value"},
		Rows: []Row{
			{Index: 0, Values: map[string]Cell{
				"Category": {Value: "A", Type: CellTypeString, RawValue: "A"},
				"Value":    {Value: 42.0, Type: CellTypeNumber, RawValue: "42"},
			}},
		},
	}

	result := table.GroupBy("Category").Aggregate(
		Sum("Value"),
		Avg("Value"),
		Min("Value"),
		Max("Value"),
		Count("Value"),
	)

	if len(result.Rows) != 1 {
		t.Fatalf("Expected 1 row, got %d", len(result.Rows))
	}

	row := result.Rows[0]

	// For single row, Sum=Avg=Min=Max=Value
	for _, col := range []string{"Sum_Value", "Avg_Value", "Min_Value", "Max_Value"} {
		cell, _ := row.Get(col)
		val, _ := cell.AsFloat()
		if val != 42 {
			t.Errorf("%s = %v, want 42", col, val)
		}
	}

	countCell, _ := row.Get("Count_Value")
	countVal, _ := countCell.AsFloat()
	if countVal != 1 {
		t.Errorf("Count = %v, want 1", countVal)
	}
}
