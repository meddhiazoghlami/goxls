// Package validation provides data validation for Excel table data.
//
// This package allows you to define validation rules and apply them to
// tables, collecting all validation errors for review.
//
// # Basic Usage
//
// Define rules and validate:
//
//	rules := []validation.ValidationRule{
//	    validation.ForColumn("Email").Required().MatchesPattern(`^[\w.-]+@[\w.-]+\.\w+$`).Build(),
//	    validation.ForColumn("Age").Range(18, 120).Build(),
//	    validation.ForColumn("Status").OneOf("active", "inactive", "pending").Build(),
//	}
//
//	result := validation.ValidateTable(table, rules)
//	if !result.Valid {
//	    for _, err := range result.Errors {
//	        fmt.Printf("Row %d, %s: %s\n", err.Row, err.Column, err.Message)
//	    }
//	}
//
// # Fluent Rule Builder
//
// Use the fluent API to build validation rules:
//
//	rule := validation.ForColumn("Email").
//	    Required().
//	    MatchesPattern(`^[\w.-]+@[\w.-]+\.\w+$`).
//	    Build()
//
// # Validation Types
//
// Available validation types:
//
//   - Required: Field cannot be empty
//   - MatchesPattern: Value must match regex pattern
//   - Range: Numeric value must be within min/max bounds
//   - OneOf: Value must be in allowed list
//   - Custom: Custom validation function
//
// # Custom Validation
//
// Create custom validation functions:
//
//	rule := validation.ValidationRule{
//	    Column: "CustomField",
//	    CustomFunc: func(cell models.Cell) error {
//	        if cell.AsString() == "invalid" {
//	            return errors.New("value cannot be 'invalid'")
//	        }
//	        return nil
//	    },
//	}
//
// # Error Grouping
//
// Group validation errors for analysis:
//
//	byColumn := result.ErrorsByColumn()  // map[string][]ValidationError
//	byRow := result.ErrorsByRow()        // map[int][]ValidationError
package validation
