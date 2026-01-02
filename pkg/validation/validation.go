package validation

import (
	"fmt"
	"regexp"
	"strings"

	"excel-lite/pkg/models"
)

// ValidationRule defines validation criteria for a column
type ValidationRule struct {
	Column        string         // Column header name to validate
	Required      bool           // If true, empty values are not allowed
	Pattern       *regexp.Regexp // Regex pattern the value must match
	MinVal        float64        // Minimum numeric value (only checked if MinValSet is true)
	MaxVal        float64        // Maximum numeric value (only checked if MaxValSet is true)
	MinValSet     bool           // Whether MinVal should be checked
	MaxValSet     bool           // Whether MaxVal should be checked
	AllowedValues []string       // List of allowed values (case-sensitive)
	CustomFunc    func(cell models.Cell) error // Custom validation function
}

// ValidationError represents a single validation failure
type ValidationError struct {
	Row     int    // Row index (0-based, relative to data rows, not including header)
	Column  string // Column name
	Value   string // The invalid value
	Message string // Description of the validation failure
}

// Error implements the error interface
func (ve ValidationError) Error() string {
	return fmt.Sprintf("row %d, column %q: %s (value: %q)", ve.Row, ve.Column, ve.Message, ve.Value)
}

// ValidationResult contains the outcome of validating a table
type ValidationResult struct {
	Valid  bool              // True if all validations passed
	Errors []ValidationError // List of all validation errors
}

// ErrorsByColumn returns validation errors grouped by column name
func (vr ValidationResult) ErrorsByColumn() map[string][]ValidationError {
	result := make(map[string][]ValidationError)
	for _, err := range vr.Errors {
		result[err.Column] = append(result[err.Column], err)
	}
	return result
}

// ErrorsByRow returns validation errors grouped by row index
func (vr ValidationResult) ErrorsByRow() map[int][]ValidationError {
	result := make(map[int][]ValidationError)
	for _, err := range vr.Errors {
		result[err.Row] = append(result[err.Row], err)
	}
	return result
}

// Validator performs validation on tables
type Validator struct {
	rules []ValidationRule
}

// NewValidator creates a new validator with the given rules
func NewValidator(rules []ValidationRule) *Validator {
	return &Validator{rules: rules}
}

// Validate validates a table against the configured rules
func (v *Validator) Validate(table *models.Table) ValidationResult {
	result := ValidationResult{Valid: true}

	if table == nil || len(table.Rows) == 0 {
		return result
	}

	for _, rule := range v.rules {
		// Check if the column exists
		columnExists := false
		for _, h := range table.Headers {
			if h == rule.Column {
				columnExists = true
				break
			}
		}
		if !columnExists {
			continue // Skip rules for non-existent columns
		}

		for rowIdx, row := range table.Rows {
			cell, exists := row.Values[rule.Column]
			if !exists {
				if rule.Required {
					result.Errors = append(result.Errors, ValidationError{
						Row:     rowIdx,
						Column:  rule.Column,
						Value:   "",
						Message: "required field is missing",
					})
					result.Valid = false
				}
				continue
			}

			errors := v.validateCell(cell, rule, rowIdx)
			if len(errors) > 0 {
				result.Errors = append(result.Errors, errors...)
				result.Valid = false
			}
		}
	}

	return result
}

// validateCell validates a single cell against a rule
func (v *Validator) validateCell(cell models.Cell, rule ValidationRule, rowIdx int) []ValidationError {
	var errors []ValidationError
	value := cell.RawValue

	// Check required
	if rule.Required && cell.IsEmpty() {
		errors = append(errors, ValidationError{
			Row:     rowIdx,
			Column:  rule.Column,
			Value:   value,
			Message: "required field is empty",
		})
		return errors // Don't continue validation if required field is empty
	}

	// Skip other validations for empty non-required fields
	if cell.IsEmpty() {
		return errors
	}

	// Check pattern
	if rule.Pattern != nil {
		if !rule.Pattern.MatchString(value) {
			errors = append(errors, ValidationError{
				Row:     rowIdx,
				Column:  rule.Column,
				Value:   value,
				Message: fmt.Sprintf("value does not match pattern %q", rule.Pattern.String()),
			})
		}
	}

	// Check numeric range
	if rule.MinValSet || rule.MaxValSet {
		if numVal, ok := cell.AsFloat(); ok {
			if rule.MinValSet && numVal < rule.MinVal {
				errors = append(errors, ValidationError{
					Row:     rowIdx,
					Column:  rule.Column,
					Value:   value,
					Message: fmt.Sprintf("value %v is less than minimum %v", numVal, rule.MinVal),
				})
			}
			if rule.MaxValSet && numVal > rule.MaxVal {
				errors = append(errors, ValidationError{
					Row:     rowIdx,
					Column:  rule.Column,
					Value:   value,
					Message: fmt.Sprintf("value %v exceeds maximum %v", numVal, rule.MaxVal),
				})
			}
		} else if rule.MinValSet || rule.MaxValSet {
			// Value is not numeric but range validation was requested
			errors = append(errors, ValidationError{
				Row:     rowIdx,
				Column:  rule.Column,
				Value:   value,
				Message: "value is not numeric but range validation was specified",
			})
		}
	}

	// Check allowed values
	if len(rule.AllowedValues) > 0 {
		found := false
		for _, allowed := range rule.AllowedValues {
			if value == allowed {
				found = true
				break
			}
		}
		if !found {
			errors = append(errors, ValidationError{
				Row:     rowIdx,
				Column:  rule.Column,
				Value:   value,
				Message: fmt.Sprintf("value not in allowed list: [%s]", strings.Join(rule.AllowedValues, ", ")),
			})
		}
	}

	// Check custom function
	if rule.CustomFunc != nil {
		if err := rule.CustomFunc(cell); err != nil {
			errors = append(errors, ValidationError{
				Row:     rowIdx,
				Column:  rule.Column,
				Value:   value,
				Message: err.Error(),
			})
		}
	}

	return errors
}

// ValidateTable is a convenience function to validate a table with given rules
func ValidateTable(table *models.Table, rules []ValidationRule) ValidationResult {
	return NewValidator(rules).Validate(table)
}

// RuleBuilder provides a fluent API for building validation rules
type RuleBuilder struct {
	rule ValidationRule
}

// ForColumn starts building a rule for a specific column
func ForColumn(name string) *RuleBuilder {
	return &RuleBuilder{rule: ValidationRule{Column: name}}
}

// Required marks the field as required
func (rb *RuleBuilder) Required() *RuleBuilder {
	rb.rule.Required = true
	return rb
}

// MatchesPattern requires the value to match a regex pattern
func (rb *RuleBuilder) MatchesPattern(pattern string) *RuleBuilder {
	rb.rule.Pattern = regexp.MustCompile(pattern)
	return rb
}

// Min sets the minimum allowed numeric value
func (rb *RuleBuilder) Min(val float64) *RuleBuilder {
	rb.rule.MinVal = val
	rb.rule.MinValSet = true
	return rb
}

// Max sets the maximum allowed numeric value
func (rb *RuleBuilder) Max(val float64) *RuleBuilder {
	rb.rule.MaxVal = val
	rb.rule.MaxValSet = true
	return rb
}

// Range sets both minimum and maximum allowed numeric values
func (rb *RuleBuilder) Range(min, max float64) *RuleBuilder {
	return rb.Min(min).Max(max)
}

// OneOf restricts the value to a set of allowed values
func (rb *RuleBuilder) OneOf(values ...string) *RuleBuilder {
	rb.rule.AllowedValues = values
	return rb
}

// Custom adds a custom validation function
func (rb *RuleBuilder) Custom(fn func(cell models.Cell) error) *RuleBuilder {
	rb.rule.CustomFunc = fn
	return rb
}

// Build returns the constructed ValidationRule
func (rb *RuleBuilder) Build() ValidationRule {
	return rb.rule
}
