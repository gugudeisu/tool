package ExcelTool

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// 流式打印
// WuLiao 根据上级物料代码,生成规则代码
func WuLiao() {
	defer fmt.Print("生成完成-----book.xlsx文件中查看")
	d := map[string]string{"02.01.001": "钢带定子71-38-大眼", "02.01.002": "钢带定子71-38-中眼"}
	f := excelize.NewFile()
	data := TidyData(d)
	//写入数据
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
	err := f.SaveAs("Book.xlsx")
	if err != nil {
		return
	}

	//close(f)
}

// 数据处理
func TidyData(d map[string]string) map[string]string {
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	data := map[string]string{}
	var i, y, z int
	for k := range d {

		//SheetNames := GetSheetName(, index)
		//数据处理
		//data[index+"1"] = d[k]

		for y = 0; y < 3; y++ {
			//十位tens place
			tensPlace := (y + 1) / 27
			//个位ones place
			onesPlace := (y + 1) % 27
			index := Slice[tensPlace] + Slice[onesPlace]
			if y == 0 {
				var dataKey string = index + strconv.Itoa(i+2)
				data[dataKey] = k
			} else if y == 1 {
				var dataKey string = index + strconv.Itoa(i+2)
				data[dataKey] = d[k]
			} else if y == 2 {

			}
		}
		// i++用于excel换行作用
		i++

		for z = 0; z < 28; z++ {
			for y = 0; y < 3; y++ {
				//十位tens place
				tensPlace := (y + 1) / 27
				//个位ones place
				onesPlace := (y + 1) % 27
				index := Slice[tensPlace] + Slice[onesPlace]
				if y == 0 {
					var dataKey string = index + strconv.Itoa(i+2)
					data[dataKey] = k + "." + convertString(z+1)
				} else if y == 1 {
					var dataKey string = index + strconv.Itoa(i+2)
					data[dataKey] = d[k]
				} else if y == 2 {
					var dataKey string = index + strconv.Itoa(i+2)
					data[dataKey] = strconv.Itoa(50 + z*10)
				}

			}
			i++
		}

		//fmt.Print(d[i], "\n")
		//Traversal(&SheetNames)

	}
	return data
}

// 转换为 001 的文本
func convertString(i int) string {
	var d string
	if i < 10 {
		d = "00" + strconv.Itoa(i)
	} else if i < 100 {
		d = "0" + strconv.Itoa(i)
	} else {
		d = strconv.Itoa(i)
	}
	return d
}

// 读取excel内容
