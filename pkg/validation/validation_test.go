package validation

import (
	"errors"
	"fmt"
	"regexp"
	"testing"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// =============================================================================
// Test Helpers
// =============================================================================

func createTestTable(headers []string, data [][]interface{}) *models.Table {
	table := &models.Table{
		Name:    "TestTable",
		Headers: headers,
		Rows:    make([]models.Row, len(data)),
	}

	for i, rowData := range data {
		row := models.Row{
			Index:  i,
			Values: make(map[string]models.Cell),
			Cells:  make([]models.Cell, len(headers)),
		}

		for j, header := range headers {
			var cell models.Cell
			if j < len(rowData) && rowData[j] != nil {
				cell = createCell(rowData[j], i, j)
			} else {
				cell = models.Cell{
					Type:     models.CellTypeEmpty,
					Row:      i,
					Col:      j,
					RawValue: "",
				}
			}
			row.Values[header] = cell
			row.Cells[j] = cell
		}

		table.Rows[i] = row
	}

	return table
}

func createCell(value interface{}, row, col int) models.Cell {
	cell := models.Cell{Row: row, Col: col}

	switch v := value.(type) {
	case string:
		cell.Value = v
		cell.RawValue = v
		if v == "" {
			cell.Type = models.CellTypeEmpty
		} else {
			cell.Type = models.CellTypeString
		}
	case int:
		cell.Value = float64(v)
		cell.RawValue = fmt.Sprintf("%d", v)
		cell.Type = models.CellTypeNumber
	case float64:
		cell.Value = v
		cell.RawValue = fmt.Sprintf("%g", v)
		cell.Type = models.CellTypeNumber
	case bool:
		cell.Value = v
		cell.RawValue = "true"
		if !v {
			cell.RawValue = "false"
		}
		cell.Type = models.CellTypeBool
	}

	return cell
}

// =============================================================================
// Basic Validation Tests
// =============================================================================

func TestNewValidator(t *testing.T) {
	rules := []ValidationRule{
		{Column: "Name", Required: true},
	}
	v := NewValidator(rules)
	if v == nil {
		t.Fatal("NewValidator() returned nil")
	}
}

func TestValidator_Validate_NilTable(t *testing.T) {
	v := NewValidator([]ValidationRule{{Column: "Name", Required: true}})
	result := v.Validate(nil)

	if !result.Valid {
		t.Error("Validate(nil) should return Valid=true")
	}
}

func TestValidator_Validate_EmptyTable(t *testing.T) {
	table := &models.Table{Headers: []string{"Name"}, Rows: []models.Row{}}
	v := NewValidator([]ValidationRule{{Column: "Name", Required: true}})
	result := v.Validate(table)

	if !result.Valid {
		t.Error("Validate() on empty table should return Valid=true")
	}
}

func TestValidator_Validate_Required(t *testing.T) {
	table := createTestTable(
		[]string{"Name", "Email"},
		[][]interface{}{
			{"Alice", "alice@example.com"},
			{"", "bob@example.com"}, // Empty name
			{"Charlie", ""},         // Empty email (not required)
		},
	)

	rules := []ValidationRule{
		{Column: "Name", Required: true},
	}

	result := NewValidator(rules).Validate(table)

	if result.Valid {
		t.Error("Expected validation to fail for empty required field")
	}

	if len(result.Errors) != 1 {
		t.Errorf("Expected 1 error, got %d", len(result.Errors))
	}

	if result.Errors[0].Row != 1 {
		t.Errorf("Expected error on row 1, got row %d", result.Errors[0].Row)
	}
}

func TestValidator_Validate_Pattern(t *testing.T) {
	table := createTestTable(
		[]string{"Email"},
		[][]interface{}{
			{"valid@example.com"},
			{"invalid-email"},
			{"another@valid.org"},
		},
	)

	rules := []ValidationRule{
		{Column: "Email", Pattern: regexp.MustCompile(`^[\w.-]+@[\w.-]+\.\w+$`)},
	}

	result := NewValidator(rules).Validate(table)

	if result.Valid {
		t.Error("Expected validation to fail for invalid email pattern")
	}

	if len(result.Errors) != 1 {
		t.Errorf("Expected 1 error, got %d", len(result.Errors))
	}

	if result.Errors[0].Row != 1 {
		t.Errorf("Expected error on row 1, got row %d", result.Errors[0].Row)
	}
}

func TestValidator_Validate_MinMax(t *testing.T) {
	table := createTestTable(
		[]string{"Age"},
		[][]interface{}{
			{float64(25)},
			{float64(15)}, // Below minimum
			{float64(150)}, // Above maximum
			{float64(50)},
		},
	)

	rules := []ValidationRule{
		{Column: "Age", MinVal: 18, MinValSet: true, MaxVal: 120, MaxValSet: true},
	}

	result := NewValidator(rules).Validate(table)

	if result.Valid {
		t.Error("Expected validation to fail for out-of-range values")
	}

	if len(result.Errors) != 2 {
		t.Errorf("Expected 2 errors, got %d", len(result.Errors))
	}
}

func TestValidator_Validate_AllowedValues(t *testing.T) {
	table := createTestTable(
		[]string{"Status"},
		[][]interface{}{
			{"active"},
			{"pending"},
			{"invalid_status"},
			{"inactive"},
		},
	)

	rules := []ValidationRule{
		{Column: "Status", AllowedValues: []string{"active", "pending", "inactive"}},
	}

	result := NewValidator(rules).Validate(table)

	if result.Valid {
		t.Error("Expected validation to fail for invalid status")
	}

	if len(result.Errors) != 1 {
		t.Errorf("Expected 1 error, got %d", len(result.Errors))
	}

	if result.Errors[0].Row != 2 {
		t.Errorf("Expected error on row 2, got row %d", result.Errors[0].Row)
	}
}

func TestValidator_Validate_CustomFunc(t *testing.T) {
	table := createTestTable(
		[]string{"Code"},
		[][]interface{}{
			{"ABC123"},
			{"abc123"}, // Should fail - not uppercase
			{"XYZ789"},
		},
	)

	rules := []ValidationRule{
		{
			Column: "Code",
			CustomFunc: func(cell models.Cell) error {
				val := cell.RawValue
				for _, r := range val {
					if r >= 'a' && r <= 'z' {
						return errors.New("code must be uppercase")
					}
				}
				return nil
			},
		},
	}

	result := NewValidator(rules).Validate(table)

	if result.Valid {
		t.Error("Expected validation to fail for lowercase code")
	}

	if len(result.Errors) != 1 {
		t.Errorf("Expected 1 error, got %d", len(result.Errors))
	}
}

func TestValidator_Validate_NonexistentColumn(t *testing.T) {
	table := createTestTable(
		[]string{"Name"},
		[][]interface{}{
			{"Alice"},
		},
	)

	rules := []ValidationRule{
		{Column: "NonExistent", Required: true},
	}

	result := NewValidator(rules).Validate(table)

	// Should pass - rule for non-existent column is skipped
	if !result.Valid {
		t.Error("Expected validation to pass for rule on non-existent column")
	}
}

func TestValidator_Validate_MultipleRules(t *testing.T) {
	table := createTestTable(
		[]string{"Name", "Age", "Status"},
		[][]interface{}{
			{"Alice", float64(25), "active"},
			{"", float64(15), "invalid"},  // Multiple failures
			{"Charlie", float64(30), "inactive"},
		},
	)

	rules := []ValidationRule{
		{Column: "Name", Required: true},
		{Column: "Age", MinVal: 18, MinValSet: true},
		{Column: "Status", AllowedValues: []string{"active", "inactive"}},
	}

	result := NewValidator(rules).Validate(table)

	if result.Valid {
		t.Error("Expected validation to fail")
	}

	// Row 1 should have 3 errors: empty name, age < 18, invalid status
	if len(result.Errors) != 3 {
		t.Errorf("Expected 3 errors, got %d", len(result.Errors))
	}
}

// =============================================================================
// ValidationResult Method Tests
// =============================================================================

func TestValidationResult_ErrorsByColumn(t *testing.T) {
	result := ValidationResult{
		Errors: []ValidationError{
			{Row: 0, Column: "Name", Message: "error1"},
			{Row: 1, Column: "Name", Message: "error2"},
			{Row: 0, Column: "Age", Message: "error3"},
		},
	}

	byColumn := result.ErrorsByColumn()

	if len(byColumn["Name"]) != 2 {
		t.Errorf("Expected 2 errors for Name column, got %d", len(byColumn["Name"]))
	}

	if len(byColumn["Age"]) != 1 {
		t.Errorf("Expected 1 error for Age column, got %d", len(byColumn["Age"]))
	}
}

func TestValidationResult_ErrorsByRow(t *testing.T) {
	result := ValidationResult{
		Errors: []ValidationError{
			{Row: 0, Column: "Name", Message: "error1"},
			{Row: 0, Column: "Age", Message: "error2"},
			{Row: 1, Column: "Name", Message: "error3"},
		},
	}

	byRow := result.ErrorsByRow()

	if len(byRow[0]) != 2 {
		t.Errorf("Expected 2 errors for row 0, got %d", len(byRow[0]))
	}

	if len(byRow[1]) != 1 {
		t.Errorf("Expected 1 error for row 1, got %d", len(byRow[1]))
	}
}

func TestValidationError_Error(t *testing.T) {
	err := ValidationError{
		Row:     5,
		Column:  "Email",
		Value:   "bad-email",
		Message: "invalid format",
	}

	expected := `row 5, column "Email": invalid format (value: "bad-email")`
	if err.Error() != expected {
		t.Errorf("Error() = %q, want %q", err.Error(), expected)
	}
}

// =============================================================================
// RuleBuilder Tests
// =============================================================================

func TestRuleBuilder_Required(t *testing.T) {
	rule := ForColumn("Name").Required().Build()

	if rule.Column != "Name" {
		t.Errorf("Column = %q, want %q", rule.Column, "Name")
	}
	if !rule.Required {
		t.Error("Required should be true")
	}
}

func TestRuleBuilder_MatchesPattern(t *testing.T) {
	rule := ForColumn("Email").MatchesPattern(`^.+@.+$`).Build()

	if rule.Pattern == nil {
		t.Fatal("Pattern should not be nil")
	}

	if !rule.Pattern.MatchString("test@example.com") {
		t.Error("Pattern should match valid email")
	}
}

func TestRuleBuilder_Range(t *testing.T) {
	rule := ForColumn("Age").Range(0, 150).Build()

	if !rule.MinValSet || rule.MinVal != 0 {
		t.Errorf("MinVal = %v (set=%v), want 0 (set=true)", rule.MinVal, rule.MinValSet)
	}
	if !rule.MaxValSet || rule.MaxVal != 150 {
		t.Errorf("MaxVal = %v (set=%v), want 150 (set=true)", rule.MaxVal, rule.MaxValSet)
	}
}

func TestRuleBuilder_Min(t *testing.T) {
	rule := ForColumn("Price").Min(0).Build()

	if !rule.MinValSet || rule.MinVal != 0 {
		t.Error("MinVal should be set to 0")
	}
	if rule.MaxValSet {
		t.Error("MaxVal should not be set")
	}
}

func TestRuleBuilder_Max(t *testing.T) {
	rule := ForColumn("Quantity").Max(1000).Build()

	if !rule.MaxValSet || rule.MaxVal != 1000 {
		t.Error("MaxVal should be set to 1000")
	}
	if rule.MinValSet {
		t.Error("MinVal should not be set")
	}
}

func TestRuleBuilder_OneOf(t *testing.T) {
	rule := ForColumn("Status").OneOf("A", "B", "C").Build()

	if len(rule.AllowedValues) != 3 {
		t.Errorf("AllowedValues length = %d, want 3", len(rule.AllowedValues))
	}
}

func TestRuleBuilder_Custom(t *testing.T) {
	customFn := func(cell models.Cell) error {
		return nil
	}
	rule := ForColumn("Code").Custom(customFn).Build()

	if rule.CustomFunc == nil {
		t.Error("CustomFunc should not be nil")
	}
}

func TestRuleBuilder_Chaining(t *testing.T) {
	rule := ForColumn("Score").
		Required().
		Range(0, 100).
		Build()

	if rule.Column != "Score" {
		t.Error("Column should be Score")
	}
	if !rule.Required {
		t.Error("Required should be true")
	}
	if !rule.MinValSet || rule.MinVal != 0 {
		t.Error("MinVal should be 0")
	}
	if !rule.MaxValSet || rule.MaxVal != 100 {
		t.Error("MaxVal should be 100")
	}
}

// =============================================================================
// Convenience Function Tests
// =============================================================================

func TestValidateTable(t *testing.T) {
	table := createTestTable(
		[]string{"Name"},
		[][]interface{}{
			{"Alice"},
			{""},
		},
	)

	rules := []ValidationRule{{Column: "Name", Required: true}}

	result := ValidateTable(table, rules)

	if result.Valid {
		t.Error("Expected validation to fail")
	}
}

// =============================================================================
// Edge Cases
// =============================================================================

func TestValidator_Validate_EmptyNonRequired(t *testing.T) {
	table := createTestTable(
		[]string{"Name", "Nickname"},
		[][]interface{}{
			{"Alice", ""},
			{"Bob", "Bobby"},
		},
	)

	rules := []ValidationRule{
		{Column: "Nickname", Pattern: regexp.MustCompile(`^[A-Z]`)}, // Only validates non-empty
	}

	result := NewValidator(rules).Validate(table)

	// Should pass - empty non-required fields skip pattern validation
	// Only Bobby should be validated, and it fails (starts with B which is uppercase - wait it matches!)
	// Let me reconsider - "Bobby" starts with B which matches [A-Z]
	if !result.Valid {
		t.Errorf("Expected validation to pass, got errors: %v", result.Errors)
	}
}

func TestValidator_Validate_NumericRangeOnNonNumeric(t *testing.T) {
	table := createTestTable(
		[]string{"Value"},
		[][]interface{}{
			{"not-a-number"},
		},
	)

	rules := []ValidationRule{
		{Column: "Value", MinVal: 0, MinValSet: true},
	}

	result := NewValidator(rules).Validate(table)

	if result.Valid {
		t.Error("Expected validation to fail for non-numeric value with range rule")
	}
}

func TestValidator_Validate_AllValidData(t *testing.T) {
	table := createTestTable(
		[]string{"Name", "Age", "Status"},
		[][]interface{}{
			{"Alice", float64(25), "active"},
			{"Bob", float64(30), "inactive"},
			{"Charlie", float64(45), "pending"},
		},
	)

	rules := []ValidationRule{
		{Column: "Name", Required: true},
		{Column: "Age", MinVal: 18, MinValSet: true, MaxVal: 100, MaxValSet: true},
		{Column: "Status", AllowedValues: []string{"active", "inactive", "pending"}},
	}

	result := NewValidator(rules).Validate(table)

	if !result.Valid {
		t.Errorf("Expected validation to pass, got errors: %v", result.Errors)
	}

	if len(result.Errors) != 0 {
		t.Errorf("Expected 0 errors, got %d", len(result.Errors))
	}
}
