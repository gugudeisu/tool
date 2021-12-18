package main

import "tool/golang"

func main() {
	d := golang.GetFileList()
	golang.Traversal(&d)
	a := "C:\\Users\\lenovo\\OneDrive\\桌面\\宇森\\物料-2021-12-16\\精工.xlsx"
	data := golang.GetSheetName(a)
	golang.Traversal(&data)
}
