package stream

// StreamOption is a functional option for configuring the stream reader
type StreamOption func(*streamOptions)

// streamOptions holds all configuration for streaming
type streamOptions struct {
	config StreamConfig
}

// WithStreamHeaders sets explicit column headers.
// When provided, these headers are used instead of reading from the first row.
// This is useful when the file has no header row or when you want to override headers.
//
// Example:
//
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamHeaders("ID", "Name", "Amount"),
//	)
func WithStreamHeaders(headers ...string) StreamOption {
	return func(o *streamOptions) {
		o.config.Headers = headers
		// If explicit headers are provided, don't read them from file
		o.config.HasHeaders = false
	}
}

// WithStreamNoHeaders indicates the file has no header row.
// Column names will be generated as Column_1, Column_2, etc.
//
// Example:
//
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamNoHeaders(),
//	)
func WithStreamNoHeaders() StreamOption {
	return func(o *streamOptions) {
		o.config.HasHeaders = false
	}
}

// WithStreamHasHeaders explicitly enables header reading from first row.
// This is the default behavior but can be used to override other options.
func WithStreamHasHeaders() StreamOption {
	return func(o *streamOptions) {
		o.config.HasHeaders = true
	}
}

// WithStreamSkipRows sets the number of rows to skip before reading headers/data.
// Use this when the file has title rows or other non-data content at the top.
//
// Example:
//
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamSkipRows(2), // Skip 2 title rows
//	)
func WithStreamSkipRows(n int) StreamOption {
	return func(o *streamOptions) {
		if n >= 0 {
			o.config.SkipRows = n
		}
	}
}

// WithStreamDateFormats sets custom date formats for type detection.
// These formats are used when inferring whether a string value is a date.
// Formats should follow Go's time package layout (e.g., "2006-01-02").
//
// Example:
//
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamDateFormats("2006-01-02", "01/02/2006", "02-Jan-2006"),
//	)
func WithStreamDateFormats(formats ...string) StreamOption {
	return func(o *streamOptions) {
		o.config.DateFormats = formats
	}
}

// WithStreamTypeDetection enables or disables automatic type detection.
// When enabled (default), cell values are parsed as numbers, booleans, dates, etc.
// When disabled, all values are returned as strings.
//
// Example:
//
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamTypeDetection(false), // All values as strings
//	)
func WithStreamTypeDetection(enabled bool) StreamOption {
	return func(o *streamOptions) {
		o.config.DetectTypes = enabled
	}
}

// WithStreamSkipEmptyRows enables or disables skipping of empty rows.
// When enabled (default), rows where all cells are empty are skipped.
// When disabled, empty rows are returned with empty cells.
//
// Example:
//
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamSkipEmptyRows(false), // Include empty rows
//	)
func WithStreamSkipEmptyRows(skip bool) StreamOption {
	return func(o *streamOptions) {
		o.config.SkipEmptyRows = skip
	}
}

// WithStreamTrimSpaces enables or disables whitespace trimming.
// When enabled (default), leading and trailing whitespace is removed from cell values.
// When disabled, whitespace is preserved.
//
// Example:
//
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamTrimSpaces(false), // Preserve whitespace
//	)
func WithStreamTrimSpaces(trim bool) StreamOption {
	return func(o *streamOptions) {
		o.config.TrimSpaces = trim
	}
}

// WithStreamConfig sets the full streaming configuration.
// This overrides all other options with the provided config.
//
// Example:
//
//	config := stream.StreamConfig{
//	    HasHeaders:    true,
//	    SkipRows:      1,
//	    DetectTypes:   true,
//	    TrimSpaces:    true,
//	    SkipEmptyRows: true,
//	}
//	sr, _ := stream.NewStreamReader("data.xlsx", "Sheet1",
//	    stream.WithStreamConfig(config),
//	)
func WithStreamConfig(config StreamConfig) StreamOption {
	return func(o *streamOptions) {
		o.config = config
	}
}
