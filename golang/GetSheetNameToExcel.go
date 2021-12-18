package golang

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

// GetSheetName 读取excel文件中所有sheet名称
func GetSheetName(p string) []string {
	var Data []string
	//f := excelize.NewFile()
	f, err := excelize.OpenFile(p)
	if err != nil {
		fmt.Print(err.Error())
		return nil
	}
	for _, sheet := range f.GetSheetList() {
		//fmt.Print(sheet)
		Data = append(Data, sheet)
		//fmt.Println(data[0])

	}
	return Data

}
