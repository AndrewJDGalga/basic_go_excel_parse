package main

import (
	"errors"
	"fmt"
	"os"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

func writeContent(fileName string, details string) {
	file, err := os.OpenFile(fileName, os.O_APPEND, 0644)
	if err != nil {
		fmt.Println("Path error!")
	}
	defer file.Close()

	_, err = file.WriteString(details)
	if err != nil {
		fmt.Println("Could not write to file")
		panic(err)
	}
	fmt.Println("Success")
}

func main() {
	filename := "{file}"
	workbook, err := xlsx.OpenFile(filename)
	section := ""

	if err != nil {
		panic(err)
	}
	sheet, ok := workbook.Sheet["{excel_tab}"]
	if !ok {
		panic(errors.New("sheet not found"))
	}

	cellProcess := func(c *xlsx.Cell) error {
		value, err := c.FormattedValue()
		substr := "{search_prefix}"
		if err != nil {
			panic(err)
		}

		if strings.Contains(value, substr) {
			section += value
			section += "\n"
		}
		return err
	}
	rowProcess := func(row *xlsx.Row) error {
		section += "','"
		return row.ForEachCell(cellProcess)
	}
	sheet.ForEachRow(rowProcess)

	writeContent("test.txt", section)
}
