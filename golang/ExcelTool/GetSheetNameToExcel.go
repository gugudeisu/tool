package ExcelTool

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

// GetSheetNameToExcel 创建一个名为sheetnames.xlsx的文件
func GetSheetNameToExcel() {
	d := GetFileList()
	f := excelize.NewFile()

	data := datalist(d)
	//写入数据
	defer fmt.Print("生成完成-----于sheetnames.xlsx文件中查看")
	for k, v := range data {
		//设置单元格的值
		err := f.SetCellValue("Sheet1", k, v)
		if err != nil {
			fmt.Print("SetCellValue, err:", err)
			return
		}
	}

	//设置默认打开的表单
	//f.SetActiveSheet(index)

	//err = SetExcelFile()
	//if err != nil {
	//	fmt.Print(err)
	//}
	err := f.SaveAs("sheetnames.xlsx")
	if err != nil {
		return
	}

	//close(f)

}

//处理data数据
func datalist(d []string) map[string]string {
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	data := map[string]string{}
	for i := range d {
		//十位tens place
		tensPlace := (i + 1) / 27
		//个位ones place
		onesPlace := (i + 1) % 27
		index := Slice[tensPlace] + Slice[onesPlace]
		//SheetNames := GetSheetName(, index)
		//数据处理
		f, err := excelize.OpenFile(d[i])
		if err != nil {
			fmt.Print("OpenFile, err:", err.Error())
			return nil
		}

		data[index+"1"] = d[i]
		for i, sheet := range f.GetSheetList() {
			var dataKey string = index + strconv.Itoa(i+2)
			data[dataKey] = sheet
		}

		//fmt.Print(d[i], "\n")
		//Traversal(&SheetNames)

	}
	return data
}

/*
// GetSheetName 读取excel文件中所有sheet名称
func GetSheetName(p string, index string) map[string]string {
	//f := excelize.NewFile()
	var data
	f, err := excelize.OpenFile(p)
	if err != nil {
		fmt.Print("OpenFile, err:", err.Error())
		return nil
	}
	for i, sheet := range f.GetSheetList() {
		var dataKey string = index + strconv.Itoa(i)
		data[dataKey] = sheet
		//data = append(data, sheet)
		//fmt.Println(data[0])

	}
	return data
}

*/

//获取各个位数上的数字
