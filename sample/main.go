package main

import (
	"fmt"

	"github.com/mususu247/oleXL"
)

func main() {
	var xl oleXL.Excel
	err := xl.Init()
	if err != nil {
		fmt.Printf("xl.Init() err: %v\n", err)
	}

	xl.CreateObject()
	xl.Visible(true)
	xl.Workbooks().Add()

	xl.Quit()
	xl.Nothing()
}
