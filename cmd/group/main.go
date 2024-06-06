package main

import (
	"fmt"
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

	alignCenter := &excelize.Alignment{
		Horizontal: "center",
		Vertical:   "center",
	}

	alicCenterStyleID, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   48,
			Family: "Consolas",
		},
		Alignment: alignCenter,
	})

	nameStyleID, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   48,
			Family: "Consolas",
		},
		Alignment: alignCenter,
	})

	_, _ = f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 5},
			{Type: "top", Color: "000000", Style: 5},
			{Type: "bottom", Color: "000000", Style: 5},
			{Type: "right", Color: "000000", Style: 5},
			// {Type: "diagonalDown", Color: "000000", Style: 5},
			// {Type: "diagonalUp", Color: "000000", Style: 5},
		},
	})

	itemStyleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   24,
			Family: "Arial",
		},
		Alignment: alignCenter,
	})
	if err != nil {
		log.Fatal(err)
	}

	row := 1
	for numData, g := range testData {
		numData++ // 從1開始

		// A
		_ = f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), numData)
		topLeftCell := fmt.Sprintf("A%d", row+1)
		_ = f.SetCellValue("Sheet1", topLeftCell, g.Name)
		_ = f.SetCellStyle("Sheet1", topLeftCell, topLeftCell, nameStyleID)
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
			cell, _ := excelize.CoordinatesToCellName(col, row)
			_ = f.SetCellValue("Sheet1", cell, idx+1) // index 1, 2, ...
			_ = f.SetCellStyle("Sheet1", cell, cell, alicCenterStyleID)

			cell, _ = excelize.CoordinatesToCellName(col, row+1)
			_ = f.SetCellValue("Sheet1", cell, string(item.Ch))
			_ = f.SetCellStyle("Sheet1", cell, cell, itemStyleID)

			// Formula
			cell, _ = excelize.CoordinatesToCellName(col, row+2)
			_ = f.SetCellFormula("Sheet1", cell, fmt.Sprintf(`=DEC2HEX("%s")`, item.Unicode))
			_ = f.SetCellStyle("Sheet1", cell, cell, itemStyleID)
		}
		row += g.Items[0].NumFields() + 1
	}

	if err = f.SaveAs("temp.output.xlsx"); err != nil {
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
