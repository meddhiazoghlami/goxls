package dateutil

import (
	"math"
	"testing"
	"time"
)

func TestExcelDateToTime(t *testing.T) {
	tests := []struct {
		name     string
		serial   float64
		expected time.Time
	}{
		{
			name:     "Excel day 1 - January 1, 1900",
			serial:   1,
			expected: time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "Excel day 2 - January 2, 1900",
			serial:   2,
			expected: time.Date(1900, 1, 2, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "February 28, 1900 (day 59)",
			serial:   59,
			expected: time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "March 1, 1900 (day 61, after leap year bug)",
			serial:   61,
			expected: time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "January 1, 2000",
			serial:   36526,
			expected: time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "January 1, 2025",
			serial:   45658,
			expected: time.Date(2025, 1, 1, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "With time component - noon",
			serial:   45658.5,
			expected: time.Date(2025, 1, 1, 12, 0, 0, 0, time.UTC),
		},
		{
			name:     "With time component - 6:00 AM",
			serial:   45658.25,
			expected: time.Date(2025, 1, 1, 6, 0, 0, 0, time.UTC),
		},
		{
			name:     "With time component - 6:00 PM",
			serial:   45658.75,
			expected: time.Date(2025, 1, 1, 18, 0, 0, 0, time.UTC),
		},
		{
			name:     "Zero returns epoch (Dec 31, 1899)",
			serial:   0,
			expected: time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC),
		},
		{
			name:     "Negative returns zero time",
			serial:   -1,
			expected: time.Time{},
		},
		{
			name:     "December 31, 2099",
			serial:   73050,
			expected: time.Date(2099, 12, 31, 0, 0, 0, 0, time.UTC),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := ExcelDateToTime(tt.serial)
			if !result.Equal(tt.expected) {
				t.Errorf("ExcelDateToTime(%v) = %v, expected %v", tt.serial, result, tt.expected)
			}
		})
	}
}

func TestTimeToExcelDate(t *testing.T) {
	tests := []struct {
		name     string
		time     time.Time
		expected float64
	}{
		{
			name:     "January 1, 1900",
			time:     time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
			expected: 1,
		},
		{
			name:     "January 2, 1900",
			time:     time.Date(1900, 1, 2, 0, 0, 0, 0, time.UTC),
			expected: 2,
		},
		{
			name:     "February 28, 1900",
			time:     time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC),
			expected: 59,
		},
		{
			name:     "March 1, 1900",
			time:     time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC),
			expected: 61,
		},
		{
			name:     "January 1, 2000",
			time:     time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC),
			expected: 36526,
		},
		{
			name:     "January 1, 2025",
			time:     time.Date(2025, 1, 1, 0, 0, 0, 0, time.UTC),
			expected: 45658,
		},
		{
			name:     "With time - noon",
			time:     time.Date(2025, 1, 1, 12, 0, 0, 0, time.UTC),
			expected: 45658.5,
		},
		{
			name:     "Zero time returns 0",
			time:     time.Time{},
			expected: 0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := TimeToExcelDate(tt.time)
			if math.Abs(result-tt.expected) > 0.0001 {
				t.Errorf("TimeToExcelDate(%v) = %v, expected %v", tt.time, result, tt.expected)
			}
		})
	}
}

func TestRoundTrip(t *testing.T) {
	// Test that converting time -> excel -> time preserves the date
	testDates := []time.Time{
		time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC),
		time.Date(1950, 6, 15, 12, 30, 0, 0, time.UTC),
		time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC),
		time.Date(2025, 7, 4, 18, 45, 0, 0, time.UTC),
		time.Date(2099, 12, 31, 23, 59, 59, 0, time.UTC),
	}

	for _, original := range testDates {
		t.Run(original.Format("2006-01-02"), func(t *testing.T) {
			serial := TimeToExcelDate(original)
			converted := ExcelDateToTime(serial)

			// Compare with second precision (nanoseconds may have rounding errors)
			if original.Unix() != converted.Unix() {
				t.Errorf("Round trip failed: %v -> %v -> %v", original, serial, converted)
			}
		})
	}
}

func TestIsExcelDateSerial(t *testing.T) {
	tests := []struct {
		name     string
		value    float64
		expected bool
	}{
		{"Valid date - day 1", 1, true},
		{"Valid date - 2025", 45658, true},
		{"Valid date - max reasonable", 109574, true},
		{"Invalid - zero", 0, false},
		{"Invalid - negative", -1, false},
		{"Invalid - too large", 200000, false},
		{"Valid - with decimal", 45658.5, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := IsExcelDateSerial(tt.value)
			if result != tt.expected {
				t.Errorf("IsExcelDateSerial(%v) = %v, expected %v", tt.value, result, tt.expected)
			}
		})
	}
}

func TestExcelDateToTimeWithLocation(t *testing.T) {
	serial := 45658.5 // Jan 1, 2025 at noon UTC

	// Test with nil location (should return UTC)
	result := ExcelDateToTimeWithLocation(serial, nil)
	if result.Location() != time.UTC {
		t.Errorf("Expected UTC location for nil input, got %v", result.Location())
	}

	// Test with specific timezone
	nyLoc, err := time.LoadLocation("America/New_York")
	if err != nil {
		t.Skip("America/New_York timezone not available")
	}

	result = ExcelDateToTimeWithLocation(serial, nyLoc)
	if result.Location() != nyLoc {
		t.Errorf("Expected America/New_York location, got %v", result.Location())
	}

	// The time should be 7:00 AM in New York (noon UTC - 5 hours)
	if result.Hour() != 7 {
		t.Errorf("Expected hour 7 in New York, got %d", result.Hour())
	}
}

func TestFormatExcelDate(t *testing.T) {
	tests := []struct {
		name     string
		serial   float64
		layout   string
		expected string
	}{
		{
			name:     "Standard date format",
			serial:   45658,
			layout:   "2006-01-02",
			expected: "2025-01-01",
		},
		{
			name:     "Full datetime format",
			serial:   45658.5,
			layout:   "2006-01-02 15:04:05",
			expected: "2025-01-01 12:00:00",
		},
		{
			name:     "US date format",
			serial:   45658,
			layout:   "01/02/2006",
			expected: "01/01/2025",
		},
		{
			name:     "Negative serial returns empty",
			serial:   -1,
			layout:   "2006-01-02",
			expected: "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := FormatExcelDate(tt.serial, tt.layout)
			if result != tt.expected {
				t.Errorf("FormatExcelDate(%v, %q) = %q, expected %q", tt.serial, tt.layout, result, tt.expected)
			}
		})
	}
}

func TestExcelLeapYearBug(t *testing.T) {
	// Excel serial 60 is the fictional February 29, 1900
	// We can't really test what it converts to since the date doesn't exist
	// But we can verify that serial 59 and 61 are correct

	feb28 := ExcelDateToTime(59)
	if feb28.Month() != 2 || feb28.Day() != 28 || feb28.Year() != 1900 {
		t.Errorf("Serial 59 should be Feb 28, 1900, got %v", feb28)
	}

	mar1 := ExcelDateToTime(61)
	if mar1.Month() != 3 || mar1.Day() != 1 || mar1.Year() != 1900 {
		t.Errorf("Serial 61 should be Mar 1, 1900, got %v", mar1)
	}
}

func TestTimeComponentPrecision(t *testing.T) {
	// Test various time components
	testCases := []struct {
		serial  float64
		hours   int
		minutes int
	}{
		{45658.0, 0, 0},          // midnight
		{45658.25, 6, 0},         // 6:00 AM
		{45658.5, 12, 0},         // noon
		{45658.75, 18, 0},        // 6:00 PM
		{45658.041666667, 1, 0},  // ~1:00 AM
		{45658.520833333, 12, 30}, // 12:30 PM
	}

	for _, tc := range testCases {
		result := ExcelDateToTime(tc.serial)
		if result.Hour() != tc.hours {
			t.Errorf("Serial %v: expected hour %d, got %d", tc.serial, tc.hours, result.Hour())
		}
		if result.Minute() != tc.minutes {
			t.Errorf("Serial %v: expected minute %d, got %d", tc.serial, tc.minutes, result.Minute())
		}
	}
}

func BenchmarkExcelDateToTime(b *testing.B) {
	for i := 0; i < b.N; i++ {
		ExcelDateToTime(45658.5)
	}
}

func BenchmarkTimeToExcelDate(b *testing.B) {
	t := time.Date(2025, 1, 1, 12, 0, 0, 0, time.UTC)
	for i := 0; i < b.N; i++ {
		TimeToExcelDate(t)
	}
}
