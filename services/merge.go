package services

import (
	"MergeExcel/models"
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"runtime"
	"strings"
)

var colnum int
var Uniquelist map[string]int

// MergeExcel 合并excel文档
func MergeExcel(path string, record []string) error {
	// 先删除已经存在的合并文件
	merge := filepath.FromSlash(path + "/merge.xlsx")
	os.Remove(merge)
	// 把1号文件另存为一份
	f, err := excelize.OpenFile(record[0])
	if err != nil {
		return err
	}
	defer f.Close()
	sheet := f.GetSheetName(f.GetActiveSheetIndex())
	cols, err := f.Cols(sheet)
	if err != nil {
		return err
	}
	cols.Next()
	col, err := cols.Rows()
	if err != nil {
		return err
	}
	// 获取所有病人的切片，防止有行号不一致的问题
	Uniquelist = make(map[string]int, 10000)
	for i, rowCell := range col {
		Uniquelist[rowCell] = i + 1
	}
	// 获取文档当前列数
	rows, err := f.Rows(sheet)
	rows.Next()
	row, err := rows.Columns()
	colnum = len(row)
	// 另存为合并文档
	err = f.SaveAs(merge)
	if err != nil {
		return err
	}
	fmt.Println("合并 -> " + merge)
	// Merge 合并文件处理
	for i := 1; i < len(record); i++ {
		err = Merge(merge, record[i])
		if err != nil {
			return err
		}
		fmt.Println("合并 -> " + record[i])
		runtime.GC()
	}
	// 重新打开merge.xlsx删除开始两列
	f, err = excelize.OpenFile(merge)
	if err != nil {
		return err
	}
	defer f.Close()
	sheet = f.GetSheetName(f.GetActiveSheetIndex())
	_ = f.RemoveCol(sheet, "A")
	_ = f.RemoveCol(sheet, "A")
	if err := f.Save(); err != nil {
		return err
	}
	return nil
}

func Merge(merge, frompath string) error {
	// 打开需要写入的文档
	if !models.FileIsExisted(merge) {
		return errors.New(merge + ",未找到合并文件")
	}
	f, err := excelize.OpenFile(merge)
	if err != nil {
		return err
	}
	defer f.Close()
	sheet := f.GetSheetName(f.GetActiveSheetIndex())
	// 打开数据来源文档
	from, err := excelize.OpenFile(frompath)
	if err != nil {
		return err
	}
	defer from.Close()
	sheet2 := from.GetSheetName(from.GetActiveSheetIndex())

	rows, err := from.Rows(sheet2)
	if err != nil {
		return err
	}

	var num int
	var nums int

	for rows.Next() {
		num++
		col, err := rows.Columns()
		if err != nil {
			return err
		}
		if num == 1 {
			nums = len(col) - 2
		}
		var rr int
		for i, colCell := range col {
			colCell = strings.Replace(colCell, "\n", "", -1)
			colCell = strings.Replace(colCell, `	`, "", -1)
			colCell = strings.Replace(colCell, `    `, "", -1)
			if num == 1 {
				// 第一行直接写入
				switch i {
				case 0, 1:
					continue
				default:
					cell, _ := excelize.CoordinatesToCellName(colnum+i-1, 1)
					err = f.SetCellValue(sheet, cell, colCell)
					if err != nil {
						return err
					}
				}

			} else {
				if len(colCell) < 1 {
					continue
				}
				switch i {
				case 0:
					rr = Uniquelist[colCell]
				case 1:
					continue
				default:
					cell, _ := excelize.CoordinatesToCellName(colnum+i-1, rr)
					err = f.SetCellValue(sheet, cell, colCell)
					if err != nil {
						return err
					}
				}
			}
		}

	}
	if err := f.Save(); err != nil {
		return err
	}
	colnum += nums
	return nil
}
