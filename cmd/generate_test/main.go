package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()

	// Sheet 1: Simple table starting at A1
	sheet1 := "Simple"
	f.SetSheetName("Sheet1", sheet1)

	// Headers
	f.SetCellValue(sheet1, "A1", "ID")
	f.SetCellValue(sheet1, "B1", "Name")
	f.SetCellValue(sheet1, "C1", "Email")
	f.SetCellValue(sheet1, "D1", "Age")

	// Data rows
	f.SetCellValue(sheet1, "A2", 1)
	f.SetCellValue(sheet1, "B2", "Alice Smith")
	f.SetCellValue(sheet1, "C2", "alice@example.com")
	f.SetCellValue(sheet1, "D2", 28)

	f.SetCellValue(sheet1, "A3", 2)
	f.SetCellValue(sheet1, "B3", "Bob Jones")
	f.SetCellValue(sheet1, "C3", "bob@example.com")
	f.SetCellValue(sheet1, "D3", 35)

	f.SetCellValue(sheet1, "A4", 3)
	f.SetCellValue(sheet1, "B4", "Charlie Brown")
	f.SetCellValue(sheet1, "C4", "charlie@example.com")
	f.SetCellValue(sheet1, "D4", 42)

	// Sheet 2: Table with offset (starts at B3)
	sheet2 := "Offset"
	f.NewSheet(sheet2)

	// Some metadata at the top
	f.SetCellValue(sheet2, "A1", "Report Generated: 2024-01-15")

	// Headers starting at B3
	f.SetCellValue(sheet2, "B3", "Product")
	f.SetCellValue(sheet2, "C3", "Category")
	f.SetCellValue(sheet2, "D3", "Price")
	f.SetCellValue(sheet2, "E3", "Quantity")

	// Data rows
	f.SetCellValue(sheet2, "B4", "Laptop")
	f.SetCellValue(sheet2, "C4", "Electronics")
	f.SetCellValue(sheet2, "D4", 999.99)
	f.SetCellValue(sheet2, "E4", 50)

	f.SetCellValue(sheet2, "B5", "Desk Chair")
	f.SetCellValue(sheet2, "C5", "Furniture")
	f.SetCellValue(sheet2, "D5", 299.50)
	f.SetCellValue(sheet2, "E5", 100)

	f.SetCellValue(sheet2, "B6", "Notebook")
	f.SetCellValue(sheet2, "C6", "Stationery")
	f.SetCellValue(sheet2, "D6", 5.99)
	f.SetCellValue(sheet2, "E6", 500)

	// Sheet 3: Multiple tables on same sheet
	sheet3 := "Multiple"
	f.NewSheet(sheet3)

	// Table 1: Top left
	f.SetCellValue(sheet3, "A1", "Department")
	f.SetCellValue(sheet3, "B1", "Budget")
	f.SetCellValue(sheet3, "A2", "Engineering")
	f.SetCellValue(sheet3, "B2", 500000)
	f.SetCellValue(sheet3, "A3", "Marketing")
	f.SetCellValue(sheet3, "B3", 200000)
	f.SetCellValue(sheet3, "A4", "Sales")
	f.SetCellValue(sheet3, "B4", 300000)

	// Table 2: Below with gap
	f.SetCellValue(sheet3, "A8", "Region")
	f.SetCellValue(sheet3, "B8", "Revenue")
	f.SetCellValue(sheet3, "C8", "Growth")
	f.SetCellValue(sheet3, "A9", "North")
	f.SetCellValue(sheet3, "B9", 1500000)
	f.SetCellValue(sheet3, "C9", "15%")
	f.SetCellValue(sheet3, "A10", "South")
	f.SetCellValue(sheet3, "B10", 1200000)
	f.SetCellValue(sheet3, "C10", "8%")
	f.SetCellValue(sheet3, "A11", "East")
	f.SetCellValue(sheet3, "B11", 900000)
	f.SetCellValue(sheet3, "C11", "12%")

	// Sheet 4: Formulas
	sheet4 := "Formulas"
	f.NewSheet(sheet4)

	// Headers
	f.SetCellValue(sheet4, "A1", "Item")
	f.SetCellValue(sheet4, "B1", "Quantity")
	f.SetCellValue(sheet4, "C1", "Price")
	f.SetCellValue(sheet4, "D1", "Total")

	// Data rows with formula in Total column
	f.SetCellValue(sheet4, "A2", "Widget A")
	f.SetCellValue(sheet4, "B2", 10)
	f.SetCellValue(sheet4, "C2", 5.50)
	f.SetCellFormula(sheet4, "D2", "B2*C2")

	f.SetCellValue(sheet4, "A3", "Widget B")
	f.SetCellValue(sheet4, "B3", 25)
	f.SetCellValue(sheet4, "C3", 3.75)
	f.SetCellFormula(sheet4, "D3", "B3*C3")

	f.SetCellValue(sheet4, "A4", "Widget C")
	f.SetCellValue(sheet4, "B4", 15)
	f.SetCellValue(sheet4, "C4", 8.00)
	f.SetCellFormula(sheet4, "D4", "B4*C4")

	// Sum row
	f.SetCellValue(sheet4, "A5", "Grand Total")
	f.SetCellFormula(sheet4, "B5", "SUM(B2:B4)")
	f.SetCellFormula(sheet4, "C5", "AVERAGE(C2:C4)")
	f.SetCellFormula(sheet4, "D5", "SUM(D2:D4)")

	// Save the file
	if err := f.SaveAs("testdata/sample.xlsx"); err != nil {
		fmt.Printf("Error saving file: %v\n", err)
		return
	}

	fmt.Println("Test file created: testdata/sample.xlsx")
}
