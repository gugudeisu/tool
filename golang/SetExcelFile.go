package golang

import (
	"github.com/xuri/excelize/v2"
)

func SetExcelFile(filename string, data map[string]string) error {
	//打开文件
	file, err := excelize.OpenFile("sheetnames.xlsx")
	if err != nil {
		return err
	}
	//获取流式写入器
	//streamWriter, err := file.NewStreamWriter("Sheet1")
	//if err != nil {
	//	return err
	//}
	//styleID, err := file.NewStyle(`{"font":{"color":"#777777"}}`)
	//err = streamWriter.SetRow("A1", []interface{}{
	//	excelize.Cell{StyleID: styleID, Value: filename}})
	//if err != nil {
	//	return err
	//}
	//for rowID := 2; rowID <= 100; rowID++ {
	//	row := make([]interface{}, 50)
	//	for colID := 0; colID < 50; colID++ {
	//		row[colID] = rand.Intn(64)
	//	}
	//	cell, _ := excelize.CoordinatesToCellName(1, rowID)
	//	if err := streamWriter.SetRow(cell, row); err != nil {
	//		fmt.Println(err)
	//	}
	//}

	for k, v := range data {
		//设置单元格的值
		err := file.SetCellValue("成绩表", k, v)
		if err != nil {
			return err
		}
	}

	//结束流式写入
	//err = streamWriter.Flush()
	//if err != nil {
	//	return err
	//}
	//保存
	err = file.Save()
	if err != nil {
		return err
	}
	return nil
}

/*
	检查sheetnames.xlsx文件是否存在
*/
func sheetnamesif(d *[]string) bool {
	for _, v := range *d {
		if v == "sheetnames.xlsx" {
			return true
		}
	}
	return false
}
