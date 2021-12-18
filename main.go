package main

import (
	"fmt"
	"tool/golang"
)

func main() {
	d := golang.GetFileList()
	for i := range d {
		data := golang.GetSheetName(d[i])
		fmt.Print(d[i], "\n")
		golang.Traversal(&data)
	}
}
