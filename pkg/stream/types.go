package stream

import (
	"strconv"
	"strings"
	"time"

	"github.com/meddhiazoghlami/goxls/pkg/models"
)

// Default date formats for type inference
var defaultDateFormats = []string{
	"2006-01-02",
	"01/02/2006",
	"02/01/2006",
	"2006/01/02",
	"Jan 2, 2006",
	"January 2, 2006",
	"02-Jan-2006",
	"2006-01-02 15:04:05",
	"01/02/2006 15:04:05",
}

// TypeInferrer handles type detection and value parsing for streaming
type TypeInferrer struct {
	dateFormats []string
}

// NewTypeInferrer creates a new type inferrer with optional custom date formats
func NewTypeInferrer(customDateFormats []string) *TypeInferrer {
	formats := defaultDateFormats
	if len(customDateFormats) > 0 {
		formats = customDateFormats
	}
	return &TypeInferrer{
		dateFormats: formats,
	}
}

// InferType determines the CellType from a raw string value
func (ti *TypeInferrer) InferType(value string) models.CellType {
	if value == "" {
		return models.CellTypeEmpty
	}

	// Check for boolean
	lower := strings.ToLower(value)
	if lower == "true" || lower == "false" {
		return models.CellTypeBool
	}

	// Check for number (including negative, decimal, scientific notation)
	if _, err := strconv.ParseFloat(value, 64); err == nil {
		return models.CellTypeNumber
	}

	// Check for date patterns
	if ti.isDateLike(value) {
		return models.CellTypeDate
	}

	return models.CellTypeString
}

// isDateLike checks if a string looks like a date
func (ti *TypeInferrer) isDateLike(value string) bool {
	for _, format := range ti.dateFormats {
		if _, err := time.Parse(format, value); err == nil {
			return true
		}
	}
	return false
}

// ParseValue converts a raw string value to the appropriate Go type
func (ti *TypeInferrer) ParseValue(value string, cellType models.CellType) interface{} {
	switch cellType {
	case models.CellTypeEmpty:
		return nil

	case models.CellTypeNumber:
		if f, err := strconv.ParseFloat(value, 64); err == nil {
			return f
		}
		return value

	case models.CellTypeBool:
		return strings.ToLower(value) == "true"

	case models.CellTypeDate:
		return ti.parseDate(value)

	default:
		return value
	}
}

// parseDate attempts to parse a date string using configured formats
func (ti *TypeInferrer) parseDate(value string) interface{} {
	for _, format := range ti.dateFormats {
		if t, err := time.Parse(format, value); err == nil {
			return t
		}
	}
	return value
}

// DateFormats returns the list of date formats used for detection
func (ti *TypeInferrer) DateFormats() []string {
	result := make([]string, len(ti.dateFormats))
	copy(result, ti.dateFormats)
	return result
}
