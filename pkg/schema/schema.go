package schema

import (
	"bytes"
	"errors"
	"fmt"
	"regexp"
	"strings"
	"unicode"

	"github.com/meddhiazoghlami/goxcel/pkg/models"
)

var (
	// ErrEmptyTable is returned when the table has no headers
	ErrEmptyTable = errors.New("table has no headers")
	// ErrEmptyStructName is returned when struct name is empty
	ErrEmptyStructName = errors.New("struct name cannot be empty")
)

// SchemaOptions configures struct generation
type SchemaOptions struct {
	// StructName is the name of the generated struct (required)
	StructName string
	// PackageName adds a "package X" declaration at the top
	PackageName string
	// ExcelTags adds `excel:"HeaderName"` struct tags (default: true)
	ExcelTags bool
	// JSONTags adds `json:"headerName"` struct tags
	JSONTags bool
	// OmitEmpty adds omitempty to tags
	OmitEmpty bool
}

// DefaultOptions returns default schema options with the given struct name
func DefaultOptions(structName string) *SchemaOptions {
	return &SchemaOptions{
		StructName: structName,
		ExcelTags:  true,
		JSONTags:   false,
		OmitEmpty:  false,
	}
}

// Generate creates a Go struct definition from a table's headers and inferred types
func Generate(table *models.Table, opts *SchemaOptions) (string, error) {
	b, err := GenerateBytes(table, opts)
	if err != nil {
		return "", err
	}
	return string(b), nil
}

// GenerateBytes creates a Go struct definition and returns it as bytes
func GenerateBytes(table *models.Table, opts *SchemaOptions) ([]byte, error) {
	if opts == nil {
		return nil, ErrEmptyStructName
	}
	if opts.StructName == "" {
		return nil, ErrEmptyStructName
	}
	if len(table.Headers) == 0 {
		return nil, ErrEmptyTable
	}

	// Analyze columns to get types
	columnStats := table.AnalyzeColumns()
	typeMap := make(map[string]models.CellType)
	for _, stat := range columnStats {
		typeMap[stat.Name] = stat.InferredType
	}

	var buf bytes.Buffer

	// Package declaration
	if opts.PackageName != "" {
		buf.WriteString(fmt.Sprintf("package %s\n\n", opts.PackageName))
	}

	// Check if we need time import
	needsTimeImport := false
	for _, header := range table.Headers {
		if cellType, ok := typeMap[header]; ok && cellType == models.CellTypeDate {
			needsTimeImport = true
			break
		}
	}

	if needsTimeImport {
		buf.WriteString("import \"time\"\n\n")
	}

	// Struct declaration
	buf.WriteString(fmt.Sprintf("type %s struct {\n", opts.StructName))

	// Track used field names to handle duplicates
	usedNames := make(map[string]int)

	// Calculate max field name length for alignment
	maxFieldLen := 0
	fieldNames := make([]string, len(table.Headers))
	for i, header := range table.Headers {
		fieldName := sanitizeFieldName(header)
		fieldName = makeUnique(fieldName, usedNames)
		fieldNames[i] = fieldName
		if len(fieldName) > maxFieldLen {
			maxFieldLen = len(fieldName)
		}
	}

	// Reset for actual generation
	usedNames = make(map[string]int)

	// Generate fields
	for i, header := range table.Headers {
		fieldName := sanitizeFieldName(header)
		fieldName = makeUnique(fieldName, usedNames)

		cellType := models.CellTypeString
		if t, ok := typeMap[header]; ok {
			cellType = t
		}
		goType := cellTypeToGoType(cellType)

		// Build struct tags
		tags := buildTags(header, opts)

		// Format with alignment
		padding := strings.Repeat(" ", maxFieldLen-len(fieldNames[i])+1)
		if tags != "" {
			buf.WriteString(fmt.Sprintf("\t%s%s%s %s\n", fieldNames[i], padding, goType, tags))
		} else {
			buf.WriteString(fmt.Sprintf("\t%s%s%s\n", fieldNames[i], padding, goType))
		}
	}

	buf.WriteString("}\n")

	return buf.Bytes(), nil
}

// cellTypeToGoType converts a CellType to its Go type string
func cellTypeToGoType(ct models.CellType) string {
	switch ct {
	case models.CellTypeNumber:
		return "float64"
	case models.CellTypeDate:
		return "time.Time"
	case models.CellTypeBool:
		return "bool"
	case models.CellTypeEmpty:
		return "interface{}"
	default: // CellTypeString, CellTypeFormula
		return "string"
	}
}

// sanitizeFieldName converts a header to a valid Go identifier
func sanitizeFieldName(header string) string {
	if header == "" {
		return "Field"
	}

	// Replace common separators with space for word detection
	header = strings.ReplaceAll(header, "_", " ")
	header = strings.ReplaceAll(header, "-", " ")

	// Remove invalid characters
	reg := regexp.MustCompile(`[^a-zA-Z0-9\s]`)
	header = reg.ReplaceAllString(header, "")

	// Split into words and capitalize each
	words := strings.Fields(header)
	if len(words) == 0 {
		return "Field"
	}

	var result strings.Builder
	for _, word := range words {
		if word == "" {
			continue
		}
		// Capitalize first letter
		runes := []rune(word)
		runes[0] = unicode.ToUpper(runes[0])
		result.WriteString(string(runes))
	}

	name := result.String()
	if name == "" {
		return "Field"
	}

	// Ensure first character is a letter
	if !unicode.IsLetter(rune(name[0])) {
		name = "Field" + name
	}

	return name
}

// makeUnique ensures field names are unique by appending a number if needed
func makeUnique(name string, used map[string]int) string {
	if count, exists := used[name]; exists {
		used[name] = count + 1
		return fmt.Sprintf("%s%d", name, count+1)
	}
	used[name] = 1
	return name
}

// buildTags creates the struct tag string
func buildTags(header string, opts *SchemaOptions) string {
	var tags []string

	if opts.ExcelTags {
		tag := fmt.Sprintf(`excel:"%s"`, header)
		if opts.OmitEmpty {
			tag = fmt.Sprintf(`excel:"%s,omitempty"`, header)
		}
		tags = append(tags, tag)
	}

	if opts.JSONTags {
		jsonName := toJSONName(header)
		tag := fmt.Sprintf(`json:"%s"`, jsonName)
		if opts.OmitEmpty {
			tag = fmt.Sprintf(`json:"%s,omitempty"`, jsonName)
		}
		tags = append(tags, tag)
	}

	if len(tags) == 0 {
		return ""
	}

	return "`" + strings.Join(tags, " ") + "`"
}

// toJSONName converts a header to a JSON-friendly name (camelCase)
func toJSONName(header string) string {
	// Replace separators with space
	header = strings.ReplaceAll(header, "_", " ")
	header = strings.ReplaceAll(header, "-", " ")

	// Remove special characters
	reg := regexp.MustCompile(`[^a-zA-Z0-9\s]`)
	header = reg.ReplaceAllString(header, "")

	words := strings.Fields(header)
	if len(words) == 0 {
		return "field"
	}

	var result strings.Builder
	for i, word := range words {
		if word == "" {
			continue
		}
		runes := []rune(strings.ToLower(word))
		if i > 0 && len(runes) > 0 {
			runes[0] = unicode.ToUpper(runes[0])
		}
		result.WriteString(string(runes))
	}

	name := result.String()
	if name == "" {
		return "field"
	}

	return name
}
