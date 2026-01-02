package main

import (
	"flag"
	"fmt"
	"os"
	"strings"

	"excel-lite/pkg/export"
	"excel-lite/pkg/models"
	"excel-lite/pkg/reader"
)

// CLI options
type options struct {
	format    string
	output    string
	sheet     string
	table     string
	columns   string
	sqlTable  string
	summary   bool
	pretty    bool
	noHeaders bool
}

func main() {
	opts := parseFlags()

	if flag.NArg() < 1 {
		printUsage()
		os.Exit(1)
	}

	filePath := flag.Arg(0)

	// Create a workbook reader
	wr := reader.NewWorkbookReader()

	// Read the file
	workbook, err := wr.ReadFile(filePath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	// Filter tables based on options
	tables := filterTables(workbook, opts)

	if len(tables) == 0 {
		fmt.Fprintln(os.Stderr, "No tables found matching criteria")
		os.Exit(1)
	}

	// Handle output based on format
	if opts.summary {
		printSummary(workbook, tables)
		return
	}

	if opts.format == "" || opts.format == "text" {
		printTextOutput(workbook, tables, opts)
		return
	}

	// Export to specified format
	output, err := exportTables(tables, opts)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Export error: %v\n", err)
		os.Exit(1)
	}

	// Write output
	if opts.output != "" {
		if err := os.WriteFile(opts.output, []byte(output), 0644); err != nil {
			fmt.Fprintf(os.Stderr, "Error writing file: %v\n", err)
			os.Exit(1)
		}
		fmt.Fprintf(os.Stderr, "Output written to: %s\n", opts.output)
	} else {
		fmt.Print(output)
	}
}

func parseFlags() options {
	var opts options
	var formatShort, outputShort, sheetShort, tableShort, columnsShort string

	flag.StringVar(&opts.format, "format", "", "Output format: json, csv, sql, text (default: text)")
	flag.StringVar(&formatShort, "f", "", "Output format (shorthand)")
	flag.StringVar(&opts.output, "output", "", "Output file path (default: stdout)")
	flag.StringVar(&outputShort, "o", "", "Output file path (shorthand)")
	flag.StringVar(&opts.sheet, "sheet", "", "Filter by sheet name")
	flag.StringVar(&sheetShort, "s", "", "Filter by sheet name (shorthand)")
	flag.StringVar(&opts.table, "table", "", "Filter by table name")
	flag.StringVar(&tableShort, "t", "", "Filter by table name (shorthand)")
	flag.StringVar(&opts.columns, "columns", "", "Comma-separated list of columns to include")
	flag.StringVar(&columnsShort, "c", "", "Columns to include (shorthand)")
	flag.StringVar(&opts.sqlTable, "sql-table", "data", "Table name for SQL output")
	flag.BoolVar(&opts.summary, "summary", false, "Show summary only")
	flag.BoolVar(&opts.pretty, "pretty", false, "Pretty print JSON output")
	flag.BoolVar(&opts.noHeaders, "no-headers", false, "Exclude headers from CSV output")

	flag.Usage = printUsage
	flag.Parse()

	// Apply shorthand flags if long form not provided
	if opts.format == "" && formatShort != "" {
		opts.format = formatShort
	}
	if opts.output == "" && outputShort != "" {
		opts.output = outputShort
	}
	if opts.sheet == "" && sheetShort != "" {
		opts.sheet = sheetShort
	}
	if opts.table == "" && tableShort != "" {
		opts.table = tableShort
	}
	if opts.columns == "" && columnsShort != "" {
		opts.columns = columnsShort
	}

	return opts
}

func printUsage() {
	fmt.Println("excel-lite - Dynamic Excel Table Reader")
	fmt.Println()
	fmt.Println("Usage: excel-lite [options] <file.xlsx>")
	fmt.Println()
	fmt.Println("Options:")
	fmt.Println("  -f, --format <format>    Output format: json, csv, sql, text (default: text)")
	fmt.Println("  -o, --output <file>      Output file path (default: stdout)")
	fmt.Println("  -s, --sheet <name>       Filter by sheet name")
	fmt.Println("  -t, --table <name>       Filter by table name")
	fmt.Println("  -c, --columns <cols>     Comma-separated columns to include")
	fmt.Println("      --sql-table <name>   Table name for SQL output (default: data)")
	fmt.Println("      --summary            Show summary information only")
	fmt.Println("      --pretty             Pretty print JSON output")
	fmt.Println("      --no-headers         Exclude headers from CSV output")
	fmt.Println()
	fmt.Println("Examples:")
	fmt.Println("  excel-lite data.xlsx")
	fmt.Println("  excel-lite data.xlsx --format=json --pretty")
	fmt.Println("  excel-lite data.xlsx -f csv -o output.csv")
	fmt.Println("  excel-lite data.xlsx --sheet=Sales --columns=Name,Amount")
	fmt.Println("  excel-lite data.xlsx -f sql --sql-table=users")
	fmt.Println("  excel-lite data.xlsx --summary")
}

func filterTables(wb *models.Workbook, opts options) []*models.Table {
	var tables []*models.Table

	for i := range wb.Sheets {
		sheet := &wb.Sheets[i]

		// Filter by sheet name
		if opts.sheet != "" && !strings.EqualFold(sheet.Name, opts.sheet) {
			continue
		}

		for j := range sheet.Tables {
			table := &sheet.Tables[j]

			// Filter by table name
			if opts.table != "" && !strings.EqualFold(table.Name, opts.table) {
				continue
			}

			tables = append(tables, table)
		}
	}

	return tables
}

func exportTables(tables []*models.Table, opts options) (string, error) {
	// Parse selected columns
	var selectedCols []string
	if opts.columns != "" {
		selectedCols = parseColumns(opts.columns)
	}

	// For single table, export directly
	if len(tables) == 1 {
		return exportTable(tables[0], opts, selectedCols)
	}

	// For multiple tables, combine output
	var results []string
	for _, table := range tables {
		result, err := exportTable(table, opts, selectedCols)
		if err != nil {
			return "", fmt.Errorf("error exporting %s: %w", table.Name, err)
		}
		results = append(results, result)
	}

	switch opts.format {
	case "json":
		// Combine JSON arrays
		return "[" + strings.Join(results, ",") + "]", nil
	case "csv":
		// Combine CSV with blank line between tables
		return strings.Join(results, "\n\n"), nil
	case "sql":
		// Combine SQL statements
		return strings.Join(results, "\n"), nil
	default:
		return strings.Join(results, "\n"), nil
	}
}

func exportTable(table *models.Table, opts options, selectedCols []string) (string, error) {
	switch opts.format {
	case "json":
		jsonOpts := export.DefaultJSONOptions()
		jsonOpts.Pretty = opts.pretty
		if len(selectedCols) > 0 {
			jsonOpts.SelectedColumns = selectedCols
		}
		exporter := export.NewJSONExporter(jsonOpts)
		return exporter.ExportString(table)

	case "csv":
		csvOpts := export.DefaultCSVOptions()
		csvOpts.IncludeHeaders = !opts.noHeaders
		if len(selectedCols) > 0 {
			csvOpts.SelectedColumns = selectedCols
		}
		exporter := export.NewCSVExporter(csvOpts)
		return exporter.ExportString(table)

	case "sql":
		sqlOpts := export.DefaultSQLOptions()
		sqlOpts.TableName = opts.sqlTable
		if len(selectedCols) > 0 {
			sqlOpts.SelectedColumns = selectedCols
		}
		exporter := export.NewSQLExporter(sqlOpts)
		return exporter.ExportString(table)

	default:
		return "", fmt.Errorf("unknown format: %s", opts.format)
	}
}

func parseColumns(cols string) []string {
	parts := strings.Split(cols, ",")
	result := make([]string, 0, len(parts))
	for _, p := range parts {
		trimmed := strings.TrimSpace(p)
		if trimmed != "" {
			result = append(result, trimmed)
		}
	}
	return result
}

func printSummary(wb *models.Workbook, tables []*models.Table) {
	fmt.Println("=== Workbook Summary ===")
	fmt.Printf("File: %s\n", wb.FilePath)
	fmt.Printf("Sheets: %d\n", len(wb.Sheets))
	fmt.Printf("Tables Found: %d\n\n", len(tables))

	for _, table := range tables {
		fmt.Printf("  %s: %d columns x %d rows\n", table.Name, len(table.Headers), len(table.Rows))
		fmt.Printf("    Headers: %v\n", table.Headers)

		// Column type analysis
		stats := table.AnalyzeColumns()
		if len(stats) > 0 {
			fmt.Println("    Column Types:")
			for _, stat := range stats {
				typeName := cellTypeName(stat.InferredType)
				fmt.Printf("      - %s: %s (%d unique values)\n", stat.Name, typeName, stat.UniqueCount)
			}
		}
		fmt.Println()
	}
}

func cellTypeName(ct models.CellType) string {
	switch ct {
	case models.CellTypeString:
		return "String"
	case models.CellTypeNumber:
		return "Number"
	case models.CellTypeDate:
		return "Date"
	case models.CellTypeBool:
		return "Boolean"
	case models.CellTypeFormula:
		return "Formula"
	default:
		return "Empty"
	}
}

func printTextOutput(wb *models.Workbook, tables []*models.Table, opts options) {
	fmt.Printf("Reading file: %s\n\n", wb.FilePath)

	fmt.Println("=== Workbook Summary ===")
	fmt.Printf("File: %s\n", wb.FilePath)
	fmt.Printf("Sheets: %d\n", len(wb.Sheets))

	totalTables := 0
	for _, sheet := range wb.Sheets {
		totalTables += len(sheet.Tables)
	}
	fmt.Printf("Total Tables Detected: %d\n\n", totalTables)

	// Parse selected columns for filtering display
	var selectedCols []string
	if opts.columns != "" {
		selectedCols = parseColumns(opts.columns)
	}

	// Group tables by sheet for display
	sheetTables := make(map[string][]*models.Table)
	for _, table := range tables {
		// Find sheet name for this table
		for _, sheet := range wb.Sheets {
			for i := range sheet.Tables {
				if &sheet.Tables[i] == table {
					sheetTables[sheet.Name] = append(sheetTables[sheet.Name], table)
					break
				}
			}
		}
	}

	for _, sheet := range wb.Sheets {
		sheetTableList, ok := sheetTables[sheet.Name]
		if !ok {
			continue
		}

		fmt.Printf("=== Sheet: %s ===\n", sheet.Name)

		for _, table := range sheetTableList {
			fmt.Printf("\n  Table: %s\n", table.Name)
			fmt.Printf("  Location: Row %d-%d, Col %d-%d\n",
				table.StartRow+1, table.EndRow+1,
				table.StartCol+1, table.EndCol+1)
			fmt.Printf("  Header Row: %d\n", table.HeaderRow+1)
			fmt.Printf("  Columns: %d\n", len(table.Headers))
			fmt.Printf("  Rows: %d\n", len(table.Rows))

			// Determine which headers to display
			displayHeaders := table.Headers
			if len(selectedCols) > 0 {
				displayHeaders = filterHeaders(table.Headers, selectedCols)
			}

			fmt.Printf("  Headers: %v\n", displayHeaders)

			// Print first few rows as sample
			if len(table.Rows) > 0 {
				fmt.Println("  Sample Data (first 3 rows):")
				maxRows := 3
				if len(table.Rows) < maxRows {
					maxRows = len(table.Rows)
				}

				for i := 0; i < maxRows; i++ {
					row := table.Rows[i]
					fmt.Printf("    Row %d: ", row.Index+1)
					first := true
					for _, header := range displayHeaders {
						if cell, ok := row.Get(header); ok {
							if !first {
								fmt.Print(", ")
							}
							fmt.Printf("%s=%q", header, cell.AsString())
							first = false
						}
					}
					fmt.Println()
				}
			}
		}
		fmt.Println()
	}
}

func filterHeaders(headers []string, selected []string) []string {
	selectedMap := make(map[string]bool)
	for _, s := range selected {
		selectedMap[strings.ToLower(s)] = true
	}

	var result []string
	for _, h := range headers {
		if selectedMap[strings.ToLower(h)] {
			result = append(result, h)
		}
	}
	return result
}
