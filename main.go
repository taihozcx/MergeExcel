package main

import (
	"MergeExcel/models"
	"MergeExcel/services"
	"fmt"
	"os"
)

func main() {
	files := models.GetCurrentAbPath()

	record := make([]string, 4)
	record[0] = files + `\1.xlsx`
	record[1] = files + `\2.xlsx`
	record[2] = files + `\3.xlsx`
	record[3] = files + `\4.xlsx`

	for i := 0; i < len(record); i++ {
		if !models.FileIsExisted(record[i]) {
			fmt.Println("需合并文件不存在->" + record[i])
			errs()
		}
	}

	err := services.MergeExcel(files, record)
	if err != nil {
		fmt.Println(err.Error())
		errs()
	} else {
		fmt.Println("合并完毕！")
	}
	os.Exit(1)
}

func errs() {
	fmt.Println("合并中止，合并过程中出现错误，按回车键结束")
	S := ""
	_, _ = fmt.Scanln(&S)
	os.Exit(0)
}
