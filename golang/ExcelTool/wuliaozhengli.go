package ExcelTool

import (
	"fmt"
	"math/rand"

	"github.com/xuri/excelize/v2"
)

func WuLiao() {

	file := excelize.NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	if err != nil {
		fmt.Println(err)
	}

	for rowID := 2; rowID <= 20000; rowID++ {
		row := make([]interface{}, 3)
		for colID := 0; colID < 3; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		if err := streamWriter.SetRow(cell, row); err != nil {
			fmt.Println(err)
		}
	}
	if err := streamWriter.Flush(); err != nil {
		fmt.Println(err)
	}
	if err := file.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

// 数据处理
func TidyData() {
	// var d []string
	// d = append(d, "123")
	// d = append(d, d)
	// fmt.Print(d)

}
