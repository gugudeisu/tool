package golang

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func GetSheetName() {
	//f := excelize.NewFile()
	data := []string{}
	f, err := excelize.OpenFile("C:\\Users\\lenovo\\OneDrive\\桌面\\宇森\\物料-2021-12-16\\精工.xlsx")
	if err != nil {
		fmt.Print(err.Error())
		return
	}
	for _, sheet := range f.GetSheetList() {
		//fmt.Print(sheet)
		data = append(data, sheet)
		fmt.Println(data[0])

	}
	print(data)
}
