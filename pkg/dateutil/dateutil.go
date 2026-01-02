package dateutil

import (
	"math"
	"time"
)

// Excel's epoch: serial 1 = January 1, 1900
// We use December 31, 1899 as the base so that adding serial days gives the correct date.
// Note: Excel incorrectly treats 1900 as a leap year (Lotus 1-2-3 bug inherited from Lotus 1-2-3).
var excelEpoch = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)

// ExcelDateToTime converts an Excel serial date number to a Go time.Time.
// Excel stores dates as the number of days since December 31, 1899 (serial 1 = Jan 1, 1900).
// Note: Excel incorrectly treats 1900 as a leap year due to a Lotus 1-2-3 bug.
// Serial number 60 is February 29, 1900 (which doesn't exist), so we adjust for dates after that.
func ExcelDateToTime(serial float64) time.Time {
	if serial < 0 {
		return time.Time{}
	}

	// Separate the integer (days) and fractional (time) parts
	days := int(serial)
	fraction := serial - float64(days)

	// Handle the 1900 leap year bug
	// Excel serial 60 = Feb 29, 1900 (doesn't exist in reality)
	// For serials > 60, we need to subtract 1 day to get the correct date
	if days > 60 {
		days--
	}

	// Calculate the time component from the fractional part
	// Use rounding to avoid floating point precision issues
	totalSeconds := fraction * 24 * 60 * 60
	totalSeconds = math.Round(totalSeconds) // Round to nearest second

	hours := int(totalSeconds) / 3600
	totalSeconds -= float64(hours * 3600)
	minutes := int(totalSeconds) / 60
	seconds := int(totalSeconds) - minutes*60

	// Add days to epoch, then set the time component
	result := excelEpoch.AddDate(0, 0, days)
	result = time.Date(result.Year(), result.Month(), result.Day(),
		hours, minutes, seconds, 0, time.UTC)

	return result
}

// TimeToExcelDate converts a Go time.Time to an Excel serial date number.
// Returns the number of days since Excel's epoch (December 31, 1899).
func TimeToExcelDate(t time.Time) float64 {
	if t.IsZero() {
		return 0
	}

	// Convert to UTC for consistent calculation
	t = t.UTC()

	// Calculate the difference in days from epoch
	duration := t.Sub(excelEpoch)
	days := duration.Hours() / 24

	// Adjust for Excel's 1900 leap year bug
	// If the date is after February 28, 1900, add 1 to account for the fake Feb 29
	cutoff := time.Date(1900, 2, 28, 23, 59, 59, 0, time.UTC)
	if t.After(cutoff) {
		days++
	}

	return days
}

// IsExcelDateSerial attempts to determine if a float64 value is likely an Excel date serial.
// Excel dates typically fall in the range of 1 (Jan 1, 1900) to ~2958465 (Dec 31, 9999).
// This is a heuristic and may not be 100% accurate.
func IsExcelDateSerial(value float64) bool {
	// Reasonable date range: 1900-01-01 to 2199-12-31
	// Serial 1 = 1900-01-01
	// Serial ~109574 = 2199-12-31
	return value >= 1 && value <= 109574
}

// ExcelDateToTimeWithLocation converts an Excel serial date to time.Time in a specific timezone.
func ExcelDateToTimeWithLocation(serial float64, loc *time.Location) time.Time {
	t := ExcelDateToTime(serial)
	if loc == nil {
		return t
	}
	return t.In(loc)
}

// FormatExcelDate converts an Excel serial date to a formatted string.
func FormatExcelDate(serial float64, layout string) string {
	t := ExcelDateToTime(serial)
	if t.IsZero() {
		return ""
	}
	return t.Format(layout)
}
