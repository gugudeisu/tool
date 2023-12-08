package ExcelTool

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"strings"
)

// 便利当前文件夹及下级汇总到同一excel
func sumexcel() {
	var a int
	fmt.Print("输入1将当前文件夹中相同sheet名称表合并\n输入2合并到同一excel\n")
	_, err := fmt.Scanf("%d", &a)
	if err != nil {
		fmt.Print("main:9 err = ", err, "\n")
		return
	}
	if a == 1 || a == 2 {
		MergeExcel(a)
	} else {
		return
	}
}

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
		if path.Ext(e) == ".xlsx" && e != "sheetnames.xlsx" && strings.Contains(e, "sheetnames.xlsx") != true {
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
