// Package schema generates Go struct definitions from Excel table data.
//
// The schema package analyzes table headers and inferred column types to
// produce Go struct code that can be used to represent the table data
// in a type-safe manner.
//
// # Basic Usage
//
// Generate a struct from a table:
//
//	code, err := schema.Generate(table, schema.DefaultOptions("Person"))
//	if err != nil {
//	    log.Fatal(err)
//	}
//	fmt.Println(code)
//
// Output:
//
//	type Person struct {
//	    Name   string  `excel:"Name"`
//	    Age    float64 `excel:"Age"`
//	    Email  string  `excel:"Email"`
//	    Active bool    `excel:"Active"`
//	}
//
// # Configuration Options
//
// Use SchemaOptions to customize the output:
//
//	opts := &schema.SchemaOptions{
//	    StructName:  "Employee",
//	    PackageName: "models",
//	    ExcelTags:   true,
//	    JSONTags:    true,
//	    OmitEmpty:   true,
//	}
//	code, err := schema.Generate(table, opts)
//
// Output:
//
//	package models
//
//	type Employee struct {
//	    Name   string  `excel:"Name,omitempty" json:"name,omitempty"`
//	    Age    float64 `excel:"Age,omitempty" json:"age,omitempty"`
//	}
//
// # Type Mapping
//
// Excel cell types are mapped to Go types as follows:
//   - CellTypeString, CellTypeFormula → string
//   - CellTypeNumber → float64
//   - CellTypeDate → time.Time
//   - CellTypeBool → bool
//   - CellTypeEmpty → interface{}
//
// # Header Sanitization
//
// Headers are automatically converted to valid Go identifiers:
//   - Spaces and special characters are removed
//   - Words are capitalized (PascalCase)
//   - Duplicate names get numeric suffixes
//   - Invalid starting characters are prefixed with "Field"
package schema
