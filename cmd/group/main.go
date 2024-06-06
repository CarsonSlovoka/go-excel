package main

import (
	"fmt"
	"github.com/CarsonSlovoka/go-excel/style"
	"github.com/xuri/excelize/v2"
	"log"
	"reflect"
)

var testData []Group

func init() {
	testData = []Group{
		{
			Name: "foo",
			Items: []Item{
				NewItem('中'),
				NewItem('文'),
			},
		},
		{
			Name: "qoo",
		},
		{
			Name: "bar",
			Items: []Item{
				NewItem('一'),
				NewItem('二'),
				NewItem('三'),
			},
		},
	}
}

func main() {
	f := excelize.NewFile()
	sM := style.NewMaker(f)

	row := 1
	var cell string
	for numData, g := range testData {
		numData++ // let start index from 1

		// A
		_ = f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), numData)
		cell, _ = excelize.CoordinatesToCellName(1, row)
		_ = f.SetCellStyle("Sheet1", cell, cell, sM.MustNewStyleID(style.AlignmentCenter, &excelize.Font{Size: 48}))
		topLeftCell := fmt.Sprintf("A%d", row+1)
		_ = f.SetCellValue("Sheet1", topLeftCell, g.Name)
		_ = f.SetCellStyle("Sheet1", topLeftCell, topLeftCell, sM.MustNewStyleID(style.AlignmentCenter, &excelize.Font{Size: 24, Family: "Consolas"}))
		if len(g.Items) == 0 {
			row += 2
			continue
		}
		// merge
		bottomRightCell := fmt.Sprintf("A%d", row+g.Items[0].NumFields())
		_ = f.MergeCell("Sheet1", topLeftCell, bottomRightCell)

		// B, C
		const dataBeginCol = 3 // C
		for idx, item := range g.Items {
			if idx == 0 {
				// left header
				_ = f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), "index")
				for i := 0; i < item.NumFields(); i++ {
					_ = f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row+i+1), item.GetFieldName(i))
				}
			}

			col := dataBeginCol + idx
			cell, _ = excelize.CoordinatesToCellName(col, row)
			_ = f.SetCellValue("Sheet1", cell, idx+1) // index 1, 2, ...
			_ = f.SetCellStyle("Sheet1", cell, cell, sM.MustNewStyleID(style.AlignmentCenter))

			cell, _ = excelize.CoordinatesToCellName(col, row+1)
			_ = f.SetCellValue("Sheet1", cell, string(item.Ch))
			_ = f.SetCellStyle("Sheet1", cell, cell, sM.MustNewStyleID(
				style.AlignmentCenter,
				&excelize.Font{Size: 18, Family: "Arial"}),
			)

			// Formula
			cell, _ = excelize.CoordinatesToCellName(col, row+2)
			_ = f.SetCellFormula("Sheet1", cell, fmt.Sprintf(`=DEC2HEX("%s")`, item.Unicode))
			_ = f.SetCellStyle("Sheet1", cell, cell, sM.MustNewStyleID(
				style.AlignmentCenter,
				&excelize.Font{Size: 18, Family: "Arial"}),
			)
		}
		row += g.Items[0].NumFields() + 1
	}

	if err := f.SaveAs("temp.output.xlsx"); err != nil {
		log.Fatalf("Error saving the Excel file: %v\n", err)
	}
}

type Group struct {
	Name  string
	Items []Item
}

type Item struct {
	Ch      rune
	Unicode string
}

func NewItem(ch rune) Item {
	return Item{
		Ch:      ch,
		Unicode: fmt.Sprintf("%d", ch),
	}
}

func (item *Item) NumFields() int {
	return reflect.ValueOf(item).Elem().NumField()
}

func (item *Item) GetFieldName(i int) string {
	t := reflect.TypeOf(*item)
	return t.Field(i).Name
}
