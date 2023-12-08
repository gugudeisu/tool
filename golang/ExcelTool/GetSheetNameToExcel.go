package ExcelTool

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

// GetSheetNameToExcel 创建一个名为sheetnames.xlsx的文件
// 获取表格sheet名
func GetSheetNameToExcel() {
	d := GetFileList()
	f := excelize.NewFile()

	data := Sheetdatalist(d, 1)
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

// 处理sheert数据
func Sheetdatalist(d []string, n int) map[string]string {
	if n != 1 && n != 2 {
		print("Sheetdatalist, err")
		return nil
	}
	data := map[string]string{}
	for i := range d {
		f, err := excelize.OpenFile(d[i])
		if err != nil {
			fmt.Print("OpenFile, err:", err.Error())
			return nil
		}

		if n == 1 {
			data[Place(i+1)+"1"] = d[i]
			for index, sheet := range f.GetSheetList() {
				var dataKey string = Place(i+1) + strconv.Itoa(index+2)
				data[dataKey] = sheet
				//fmt.Print(sheet, "|", Place(i+1)+strconv.Itoa(i+2), "|", i, "\n")
			}
		}
		if n == 2 {
			for _, sheet := range f.GetSheetList() {
				data[sheet] = sheet
				//print(sheet)
				//fmt.Print(sheet, "|", Place(i+1)+strconv.Itoa(i+2), "|", i, "\n")
			}
		}

		//fmt.Print(Place(i+1), "\n")

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
