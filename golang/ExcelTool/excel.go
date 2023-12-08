package ExcelTool

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

// Excelcharuhang 用于将excel横向内容,竖向展开 插入固定内容的行
func Excelcharuhang() {
	f := excelize.NewFile()
	data := charu()

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

// 用于将excel横向内容,竖向产开
func zhankai() map[string]string {
	data := map[string]string{}

	f, err := excelize.OpenFile("物料_缺少工艺路线.xlsx")
	if err != nil {
		fmt.Print("OpenFile, err:", err.Error())
		return nil
	}
	//sumList := 2
	Column := 1

	for h := 0; h < 1260; h++ {
		for l := 4; l < 12; l++ {
			cell, err := f.GetCellValue("Sheet1", Place(l+2)+strconv.Itoa(h+1))
			if err != nil {
				fmt.Println(err, "读取失败", "|", f.GetSheetList())
				return nil
			}

			wl, err := f.GetCellValue("Sheet1", "A"+strconv.Itoa(h+1))
			if err != nil {
				fmt.Println(err, "读取失败", "|", f.GetSheetList())
				return nil
			}

			data["A"+strconv.Itoa(Column+1)] = wl
			data["B"+strconv.Itoa(Column+1)] = cell
			fmt.Print(cell + "\n")
			Column++
		}

	}

	//fmt.Print(cell, "\n")

	return data

}

// 插入固定内容的行
func charu() map[string]string {
	data := map[string]string{}

	f, err := excelize.OpenFile("工艺路线_压轴动平衡.xlsx")
	if err != nil {
		fmt.Print("OpenFile, err:", err.Error())
		return nil
	}
	//sumList := 2
	//用于换行, 增加行数
	Column := 1
	text1 := [18]string{"100", "大洋测试", "WC000001", "cs", "84", "压轴", "压轴", "OCC000001", "免检厂内汇报", "26", "台",
		"与工作中心一致",
		"分钟",
		"分钟",
		"工序开工",
		"无",
		"手动",
		"严格控制"}

	text2 := [18]string{"100", "大洋测试", "WC000001", "cs", "85", "动平衡", "动平衡", "OCC000001", "免检厂内汇报", "26", "台",
		"与工作中心一致",
		"分钟",
		"分钟",
		"工序开工",
		"无",
		"手动",
		"严格控制"}
	/***
	h:excel起始行
	l:excel起始列
	用于读取excel内容
	***/
	for h := 0; h < 1846; h++ {
		for l := 0; l < 47; l++ {
			cell, err := f.GetCellValue("工艺路线#单据头(FBillHead)", Place(l+1)+strconv.Itoa(h+1))
			if err != nil {
				fmt.Println(err, "读取失败", "|", f.GetSheetList())
				return nil
			}

			if cell != "" {
				data[Place(l+1)+strconv.Itoa(h+Column)] = cell
			}

			//fmt.Print(cell + "\n")
			//判断该单据内是否有压轴
			//当前单元格为U列 , 且单元格内容不为空, 插入数值 并下移一行, 只改变行数 不改变原有列
			if Place(l+1) == "U" && cell != "" {

				for i := 0; i < len(text1); i++ {
					data[Place(l+5+i)+strconv.Itoa(h+Column)] = text1[i]
				}
				Column++
				for i := 0; i < len(text2); i++ {
					data[Place(l+5+i)+strconv.Itoa(h+Column)] = text2[i]
				}
				Column++
			}

		}
	}
	//fmt.Print(cell, "\n")

	return data

}
