package ExcelTool

import (
	"io/ioutil"
	"log"
	"os"
	"path"
)

// GetFileList 读取当前路径下的文件名
func GetFileList() []string {
	var Data []string
	pwd, _ := os.Getwd() //获取当前目录
	//获取文件或目录相关信息
	fileInfoList, err := ioutil.ReadDir(pwd)
	if err != nil {
		log.Fatal("ReadDir, err:", err)
	}
	for i := range fileInfoList {
		e := fileInfoList[i].Name()
		if path.Ext(e) == ".xlsx" && e != "sheetnames.xlsx" {
			Data = append(Data, fileInfoList[i].Name())
		}
	}
	return Data
}

// Traversal 遍历数组
func Traversal(d *[]string) {
	for _, v := range *d {
		println(v)
	}
}
