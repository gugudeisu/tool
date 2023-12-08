package ExcelTool

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

// 合并excel内容
func MergeExcel(n int) {
	d := GetFileList()

	sheetName := Sheetdatalist(d, 2)
	//fmt.Print(d)
	//fmt.Print("\n", sheetName)

	if n == 1 {
		for _, v := range sheetName {
			f := excelize.NewFile()
			data := datalist(d, v)
			//fmt.Print(data)
			//fmt.Print("\n")
			for k, v := range data {
				//设置单元格的值
				err := f.SetCellValue("Sheet1", k, v)
				if err != nil {
					fmt.Print("SetCellValue, err:", err, k, v)
					return
				}

				//fmt.Print(k, v, "\n")
			}
			err := f.SaveAs(v + "sheetnames.xlsx")
			if err != nil {
				return
			}

		}
	} else if n == 2 {
		f := excelize.NewFile()
		data := dataListSum(d)

		//fmt.Print(data)
		for k, v := range data {
			//设置单元格的值
			err := f.SetCellValue("Sheet1", k, v)
			if err != nil {
				fmt.Print("SetCellValue, err:", err, k, v)
				return
			}

			//fmt.Print(k, v, "\n")
		}
		err := f.SaveAs("sheetnames.xlsx")
		if err != nil {
			return
		}

	}

	//defer fmt.Print("生成完成-----于sheetnames.xlsx文件中查看")

}

// 处理data ,当 sheetname 不同时不拆分不同表
func dataListSum(d []string) map[string]string {
	data := map[string]string{}
	sumList := 1
	for i := range d {
		f, err := excelize.OpenFile(d[i])
		if err != nil {
			fmt.Print("OpenFile, err:", err.Error())
			return nil
		}
		for _, v := range f.GetSheetList() {
			Column := 0
			data[Place(1)+strconv.Itoa(sumList+1)] = d[i]
			data[Place(2)+strconv.Itoa(sumList+1)] = v
			for row := 1; row <= 28; row++ {
				for l := 0; l < 11; l++ {
					cell, err := f.GetCellValue(v, Place(l+1)+strconv.Itoa(row))
					if err != nil {
						fmt.Println(err, "读取失败", "|", f.GetSheetList(), "|", "l=", l+1, "row = ", row, Place(l+1)+strconv.Itoa(row))
						return nil
					}
					//data[Place(l)+strconv.Itoa(row)] = cell
					//data[Place(Column+1)+strconv.Itoa(i+1)+"|"+Place(l+1)+strconv.Itoa(row)+"|"+strconv.Itoa(Column)] = cell

					data[Place(Column+3)+strconv.Itoa(sumList+1)] = cell
					data[Place(Column+3)+strconv.Itoa(1)] = Place(l+1) + strconv.Itoa(row)
					//fmt.Print(cell, "\n")
					Column++
				}

			}

			sumList++
		}
		//行标识

		//data[index+"1"] = d[i]
		//for i, sheet := range f.GetSheetList() {
		//	var dataKey string = index + strconv.Itoa(i+2)
		//	data[dataKey] = sheet
		//}

		//fmt.Print(d[i], "\n")
		//Traversal(&SheetNames)

	}
	return data
}

// 处理data ,当 sheetname 不同时拆分不同表
func datalist(d []string, sheetName string) map[string]string {
	data := map[string]string{}
	for i := range d {
		f, err := excelize.OpenFile(d[i])
		if err != nil {
			fmt.Print("OpenFile, err:", err.Error())
			return nil
		}
		for _, v := range f.GetSheetList() {
			if v == sheetName {
				Column := 0
				for row := 1; row <= 28; row++ {
					for l := 0; l < 11; l++ {
						cell, err := f.GetCellValue(sheetName, Place(l+1)+strconv.Itoa(row))
						if err != nil {
							fmt.Println(err, "读取失败", "|", f.GetSheetList(), "|", "l=", l+1, "row = ", row, Place(l+1)+strconv.Itoa(row))
							return nil
						}
						//data[Place(l)+strconv.Itoa(row)] = cell
						//data[Place(Column+1)+strconv.Itoa(i+1)+"|"+Place(l+1)+strconv.Itoa(row)+"|"+strconv.Itoa(Column)] = cell
						data[Place(Column+1)+strconv.Itoa(i+1)] = cell
						Column++
					}

				}

			}
		}
		//行标识

		//data[index+"1"] = d[i]
		//for i, sheet := range f.GetSheetList() {
		//	var dataKey string = index + strconv.Itoa(i+2)
		//	data[dataKey] = sheet
		//}

		//fmt.Print(d[i], "\n")
		//Traversal(&SheetNames)

	}
	return data
}

func Place1(i int) string {
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	//十位tens place
	tensPlace := (i + 1) / 27
	//个位ones place
	onesPlace := (i + 1) % 27
	index := Slice[tensPlace] + Slice[onesPlace]
	//SheetNames := GetSheetName(, index)
	return index

}

var ExcelChar = []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func Place(num int) string {
	var cols string
	v := num
	for v > 0 {
		k := v % 26
		if k == 0 {
			k = 26
		}
		v = (v - k) / 26
		cols = ExcelChar[k] + cols
	}
	return cols
}
