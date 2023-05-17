package excel

import (
	"log"

	"github.com/xuri/excelize/v2"
)

type NewExcel struct {
	FileStream  *excelize.File
	SheetValues SheetValues
	Error       error
}

type ExcelClient interface {
	ReadAllSheetsCells() SheetValues
	ReadRow(sheet string) ([][]string, error)
	ReadColumnHeader(sheet string) (ColumnInfo, error)
	FindHyperLinkCells(sheet string) error
}

type SheetValues struct {
	Sheets     []string
	Rows       map[string][][]string
	ColumnInfo ColumnInfo
}

type ColumnInfo struct {
	Columns []string
	Index   int
}

func NewExcelEditor(path string) NewExcel {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return NewExcel{FileStream: nil, Error: err}
	}
	return NewExcel{FileStream: f, Error: nil}
}

func (e *NewExcel) ReadAllSheetsCells() SheetValues {
	// mySheetValues := make(map[string][][]string)
	for _, sheet := range e.FileStream.GetSheetMap() {
		rows, err := e.FileStream.GetRows(sheet)
		if err != nil {
			log.Fatal("failed to read sheet", err)
		}
		makeRow := make(map[string][][]string)
		makeRow[sheet] = rows
		e.SheetValues.Sheets = append(e.SheetValues.Sheets, sheet)
		e.SheetValues.Rows = makeRow
	}
	return e.SheetValues
}

func (e *NewExcel) ReadRow(sheet string) ([][]string, error) {
	rows, err := e.FileStream.GetRows(sheet)
	if err != nil {
		log.Println("failed to read rows", err)
		return nil, err
	}
	return rows, nil
}

func (e *NewExcel) ReadColumnHeader(sheet string) (ColumnInfo, error) {
	columnInfo := ColumnInfo{}
	rows, err := e.FileStream.GetRows(sheet)
	if err != nil {
		log.Println("failed to read rows", err)
		return columnInfo, err
	}

	for idx, row := range rows {
		if len(row) > 0 && row[0] == "메뉴명" {
			columnInfo.Columns = row
			columnInfo.Index = idx
			break
		}
	}

	return columnInfo, nil
}

func (e *NewExcel) FindHyperLinkCells(sheet string) error {
	columns, err := e.ReadColumnHeader(sheet)
	if err != nil {
		return err
	}

	log.Println(excelize.CellNameToCoordinates(columns.Columns[1]))
	return nil
}
