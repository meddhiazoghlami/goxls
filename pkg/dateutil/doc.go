// Package dateutil provides utilities for converting between Excel serial dates
// and Go time.Time values.
//
// Excel stores dates as serial numbers where 1 represents January 1, 1900.
// This package handles the conversion including the infamous 1900 leap year bug
// that Excel inherited from Lotus 1-2-3.
//
// # Basic Conversion
//
// Convert Excel serial date to Go time:
//
//	t := dateutil.ExcelDateToTime(45658)  // January 1, 2025
//
// Convert Go time to Excel serial:
//
//	serial := dateutil.TimeToExcelDate(time.Now())
//
// # Formatting
//
// Convert and format in one call:
//
//	formatted := dateutil.FormatExcelDate(45658, "2006-01-02")
//	// Output: "2025-01-01"
//
// # Timezone Support
//
// Convert with specific timezone:
//
//	loc, _ := time.LoadLocation("America/New_York")
//	t := dateutil.ExcelDateToTimeWithLocation(45658, loc)
//
// # Checking Values
//
// Check if a value is likely an Excel date serial:
//
//	if dateutil.IsExcelDateSerial(45658) {
//	    // Value is in reasonable date range
//	}
//
// # The 1900 Leap Year Bug
//
// Excel incorrectly treats 1900 as a leap year (February 29, 1900 exists in Excel
// but not in reality). This package handles this quirk automatically, ensuring
// dates before and after March 1, 1900 are converted correctly.
package dateutil
