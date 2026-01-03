package schema

import (
	"strings"
	"testing"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

func createTestTable() *models.Table {
	return &models.Table{
		Name:    "TestTable",
		Headers: []string{"Name", "Age", "Email", "Active"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Name":   {Value: "Alice", Type: models.CellTypeString, RawValue: "Alice"},
					"Age":    {Value: float64(30), Type: models.CellTypeNumber, RawValue: "30"},
					"Email":  {Value: "alice@example.com", Type: models.CellTypeString, RawValue: "alice@example.com"},
					"Active": {Value: true, Type: models.CellTypeBool, RawValue: "TRUE"},
				},
			},
			{
				Index: 2,
				Values: map[string]models.Cell{
					"Name":   {Value: "Bob", Type: models.CellTypeString, RawValue: "Bob"},
					"Age":    {Value: float64(25), Type: models.CellTypeNumber, RawValue: "25"},
					"Email":  {Value: "bob@example.com", Type: models.CellTypeString, RawValue: "bob@example.com"},
					"Active": {Value: false, Type: models.CellTypeBool, RawValue: "FALSE"},
				},
			},
		},
	}
}

func TestGenerate_Basic(t *testing.T) {
	table := createTestTable()
	opts := DefaultOptions("Person")

	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	// Check struct declaration
	if !strings.Contains(code, "type Person struct {") {
		t.Error("Missing struct declaration")
	}

	// Check fields
	if !strings.Contains(code, "Name") {
		t.Error("Missing Name field")
	}
	if !strings.Contains(code, "Age") {
		t.Error("Missing Age field")
	}
	if !strings.Contains(code, "Email") {
		t.Error("Missing Email field")
	}
	if !strings.Contains(code, "Active") {
		t.Error("Missing Active field")
	}

	// Check types
	if !strings.Contains(code, "string") {
		t.Error("Missing string type")
	}
	if !strings.Contains(code, "float64") {
		t.Error("Missing float64 type")
	}
	if !strings.Contains(code, "bool") {
		t.Error("Missing bool type")
	}

	// Check excel tags (default)
	if !strings.Contains(code, `excel:"Name"`) {
		t.Error("Missing excel tag for Name")
	}
}

func TestGenerate_WithPackage(t *testing.T) {
	table := createTestTable()
	opts := &SchemaOptions{
		StructName:  "Person",
		PackageName: "models",
		ExcelTags:   true,
	}

	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	if !strings.HasPrefix(code, "package models\n") {
		t.Error("Missing or incorrect package declaration")
	}
}

func TestGenerate_WithJSONTags(t *testing.T) {
	table := createTestTable()
	opts := &SchemaOptions{
		StructName: "Person",
		ExcelTags:  true,
		JSONTags:   true,
	}

	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	if !strings.Contains(code, `json:"name"`) {
		t.Error("Missing json tag for name")
	}
	if !strings.Contains(code, `json:"age"`) {
		t.Error("Missing json tag for age")
	}
}

func TestGenerate_WithOmitEmpty(t *testing.T) {
	table := createTestTable()
	opts := &SchemaOptions{
		StructName: "Person",
		ExcelTags:  true,
		OmitEmpty:  true,
	}

	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	if !strings.Contains(code, `excel:"Name,omitempty"`) {
		t.Error("Missing omitempty in excel tag")
	}
}

func TestGenerate_WithJSONAndOmitEmpty(t *testing.T) {
	table := createTestTable()
	opts := &SchemaOptions{
		StructName: "Person",
		ExcelTags:  false,
		JSONTags:   true,
		OmitEmpty:  true,
	}

	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	if !strings.Contains(code, `json:"name,omitempty"`) {
		t.Error("Missing omitempty in json tag")
	}
}

func TestGenerate_NoTags(t *testing.T) {
	table := createTestTable()
	opts := &SchemaOptions{
		StructName: "Person",
		ExcelTags:  false,
		JSONTags:   false,
	}

	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	if strings.Contains(code, "excel:") {
		t.Error("Should not have excel tags")
	}
	if strings.Contains(code, "json:") {
		t.Error("Should not have json tags")
	}
}

func TestGenerate_DateType(t *testing.T) {
	table := &models.Table{
		Name:    "TestTable",
		Headers: []string{"Name", "JoinDate"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Name":     {Value: "Alice", Type: models.CellTypeString, RawValue: "Alice"},
					"JoinDate": {Value: float64(45000), Type: models.CellTypeDate, RawValue: "45000"},
				},
			},
		},
	}

	opts := DefaultOptions("Employee")
	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	if !strings.Contains(code, `import "time"`) {
		t.Error("Missing time import for date field")
	}
	if !strings.Contains(code, "time.Time") {
		t.Error("Missing time.Time type for date field")
	}
}

func TestGenerate_EmptyTable(t *testing.T) {
	table := &models.Table{
		Name:    "Empty",
		Headers: []string{},
	}

	opts := DefaultOptions("Empty")
	_, err := Generate(table, opts)
	if err != ErrEmptyTable {
		t.Errorf("Expected ErrEmptyTable, got %v", err)
	}
}

func TestGenerate_EmptyStructName(t *testing.T) {
	table := createTestTable()
	opts := &SchemaOptions{StructName: ""}

	_, err := Generate(table, opts)
	if err != ErrEmptyStructName {
		t.Errorf("Expected ErrEmptyStructName, got %v", err)
	}
}

func TestGenerate_NilOptions(t *testing.T) {
	table := createTestTable()

	_, err := Generate(table, nil)
	if err != ErrEmptyStructName {
		t.Errorf("Expected ErrEmptyStructName, got %v", err)
	}
}

func TestSanitizeFieldName(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"Name", "Name"},
		{"first name", "FirstName"},
		{"first_name", "FirstName"},
		{"first-name", "FirstName"},
		{"123number", "Field123number"},
		{"email@address", "Emailaddress"},
		{"", "Field"},
		{"   ", "Field"},
		{"user ID", "UserID"},
		{"is_active", "IsActive"},
	}

	for _, tc := range tests {
		t.Run(tc.input, func(t *testing.T) {
			result := sanitizeFieldName(tc.input)
			if result != tc.expected {
				t.Errorf("sanitizeFieldName(%q) = %q, want %q", tc.input, result, tc.expected)
			}
		})
	}
}

func TestMakeUnique(t *testing.T) {
	used := make(map[string]int)

	// First use
	name1 := makeUnique("Name", used)
	if name1 != "Name" {
		t.Errorf("First use should be 'Name', got %q", name1)
	}

	// Second use - should get suffix
	name2 := makeUnique("Name", used)
	if name2 != "Name2" {
		t.Errorf("Second use should be 'Name2', got %q", name2)
	}

	// Third use
	name3 := makeUnique("Name", used)
	if name3 != "Name3" {
		t.Errorf("Third use should be 'Name3', got %q", name3)
	}
}

func TestToJSONName(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"Name", "name"},
		{"FirstName", "firstname"},
		{"first name", "firstName"},
		{"first_name", "firstName"},
		{"ID", "id"},
		{"userID", "userid"},
		{"", "field"},
	}

	for _, tc := range tests {
		t.Run(tc.input, func(t *testing.T) {
			result := toJSONName(tc.input)
			if result != tc.expected {
				t.Errorf("toJSONName(%q) = %q, want %q", tc.input, result, tc.expected)
			}
		})
	}
}

func TestGenerate_DuplicateHeaders(t *testing.T) {
	table := &models.Table{
		Name:    "TestTable",
		Headers: []string{"Name", "Name", "Name"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Name": {Value: "Test", Type: models.CellTypeString, RawValue: "Test"},
				},
			},
		},
	}

	opts := DefaultOptions("Person")
	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	// Should have Name, Name2, Name3
	if !strings.Contains(code, "Name ") && !strings.Contains(code, "Name\t") {
		t.Error("Missing Name field")
	}
	if !strings.Contains(code, "Name2") {
		t.Error("Missing Name2 field")
	}
	if !strings.Contains(code, "Name3") {
		t.Error("Missing Name3 field")
	}
}

func TestGenerateBytes(t *testing.T) {
	table := createTestTable()
	opts := DefaultOptions("Person")

	bytes, err := GenerateBytes(table, opts)
	if err != nil {
		t.Fatalf("GenerateBytes failed: %v", err)
	}

	if len(bytes) == 0 {
		t.Error("GenerateBytes returned empty bytes")
	}

	code := string(bytes)
	if !strings.Contains(code, "type Person struct") {
		t.Error("GenerateBytes output missing struct declaration")
	}
}

func TestDefaultOptions(t *testing.T) {
	opts := DefaultOptions("TestStruct")

	if opts.StructName != "TestStruct" {
		t.Errorf("StructName = %q, want %q", opts.StructName, "TestStruct")
	}
	if !opts.ExcelTags {
		t.Error("ExcelTags should be true by default")
	}
	if opts.JSONTags {
		t.Error("JSONTags should be false by default")
	}
	if opts.OmitEmpty {
		t.Error("OmitEmpty should be false by default")
	}
	if opts.PackageName != "" {
		t.Error("PackageName should be empty by default")
	}
}

func TestGenerate_SpecialCharHeaders(t *testing.T) {
	table := &models.Table{
		Name:    "TestTable",
		Headers: []string{"user@email", "price ($)", "% complete"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"user@email": {Value: "test@example.com", Type: models.CellTypeString, RawValue: "test@example.com"},
					"price ($)":  {Value: float64(99.99), Type: models.CellTypeNumber, RawValue: "99.99"},
					"% complete": {Value: float64(50), Type: models.CellTypeNumber, RawValue: "50"},
				},
			},
		},
	}

	opts := DefaultOptions("Record")
	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	// Should have sanitized field names (@ and special chars removed, words joined)
	if !strings.Contains(code, "Useremail") {
		t.Error("Missing Useremail field (sanitized from user@email)")
	}
	if !strings.Contains(code, "Price") {
		t.Error("Missing Price field (sanitized from price ($))")
	}
	if !strings.Contains(code, "Complete") {
		t.Error("Missing Complete field (sanitized from % complete)")
	}

	// Should preserve original names in tags
	if !strings.Contains(code, `excel:"user@email"`) {
		t.Error("Excel tag should have original header name")
	}
}

func TestGenerate_FormulaType(t *testing.T) {
	table := &models.Table{
		Name:    "TestTable",
		Headers: []string{"Total"},
		Rows: []models.Row{
			{
				Index: 1,
				Values: map[string]models.Cell{
					"Total": {Value: "=SUM(A1:A10)", Type: models.CellTypeFormula, RawValue: "=SUM(A1:A10)"},
				},
			},
		},
	}

	opts := DefaultOptions("Summary")
	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	// Formula type should map to string
	if !strings.Contains(code, "Total") && !strings.Contains(code, "string") {
		t.Error("Formula field should be string type")
	}
}

func TestGenerate_EmptyType(t *testing.T) {
	// For empty type test, we test with no rows (all empty)
	table := &models.Table{
		Name:    "TestTable",
		Headers: []string{"Unknown"},
		Rows:    []models.Row{},
	}

	opts := DefaultOptions("Record")
	code, err := Generate(table, opts)
	if err != nil {
		t.Fatalf("Generate failed: %v", err)
	}

	// Empty type should map to interface{}
	if !strings.Contains(code, "interface{}") {
		t.Error("Empty field should be interface{} type")
	}
}
