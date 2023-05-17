package main

import (
	"flag"
	"log"
	"office-automation/common"
	"office-automation/excel"
)

func main() {

	flags := common.CommandFlag{}

	flags.Filename = flag.String("path", "", "input target file path")
	flags.SheetName = flag.String("sheet", "", "input sheet name")
	flag.Parse()

	e := excel.NewExcelEditor(*flags.Filename)
	if e.Error != nil {
		log.Fatal("failed to open excel", e.Error)
	}
	defer e.FileStream.Close()

	log.Println(e.FindHyperLinkCells("Sheet1"))
}
