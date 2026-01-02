// Example: Export to SQL
//
// This example demonstrates exporting Excel data to SQL INSERT statements
// with support for different database dialects.
//
// Run: go run main.go
package main

import (
	"fmt"
	"log"

	"github.com/meddhiazoghlami/goxcel"
	"github.com/meddhiazoghlami/goxcel/pkg/export"
)

func main() {
	// Read Excel file
	workbook, err := goxcel.ReadFile("../../testdata/sample.xlsx")
	if err != nil {
		log.Fatalf("Failed to read file: %v", err)
	}

	if len(workbook.Sheets) == 0 || len(workbook.Sheets[0].Tables) == 0 {
		log.Fatal("No tables found in workbook")
	}

	table := &workbook.Sheets[0].Tables[0]

	// Method 1: Simple SQL INSERT statements
	fmt.Println("=== Basic SQL INSERT ===")
	sql, err := goxcel.ToSQL(table, "my_table")
	if err != nil {
		log.Fatalf("ToSQL failed: %v", err)
	}
	fmt.Println(sql)

	// Method 2: SQL with CREATE TABLE statement
	fmt.Println("=== SQL with CREATE TABLE ===")
	sqlCreate, err := goxcel.ToSQLWithCreate(table, "my_table")
	if err != nil {
		log.Fatalf("ToSQLWithCreate failed: %v", err)
	}
	fmt.Println(sqlCreate)

	// Method 3: PostgreSQL dialect
	fmt.Println("=== PostgreSQL Dialect ===")
	pgOpts := export.DefaultSQLOptions()
	pgOpts.TableName = "employees"
	pgOpts.Dialect = export.DialectPostgreSQL
	pgOpts.CreateTable = true

	pgExporter := export.NewSQLExporter(pgOpts)
	pgSQL, err := pgExporter.ExportString(table)
	if err != nil {
		log.Fatalf("PostgreSQL export failed: %v", err)
	}
	fmt.Println(pgSQL)

	// Method 4: MySQL dialect with batching
	fmt.Println("=== MySQL Dialect (Batched) ===")
	mysqlOpts := export.DefaultSQLOptions()
	mysqlOpts.TableName = "users"
	mysqlOpts.Dialect = export.DialectMySQL
	mysqlOpts.BatchSize = 2 // Small batch for demo

	mysqlExporter := export.NewSQLExporter(mysqlOpts)
	mysqlSQL, err := mysqlExporter.ExportString(table)
	if err != nil {
		log.Fatalf("MySQL export failed: %v", err)
	}
	fmt.Println(mysqlSQL)
}
