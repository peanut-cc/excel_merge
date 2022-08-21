package main

import (
	"fmt"
	"github.com/peanut-cc/excel_merge/utils/uuid"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"log"
	"time"
)

// GetAllFiles 获取目录下的所有文件的绝对路径
func GetAllFiles(pathName string) ([]string, error) {
	var result []string
	rd, err := ioutil.ReadDir(pathName)
	if err != nil {
		log.Printf("读取路径[%s]下文件错误", pathName)
		return nil, err
	}
	for _, fi := range rd {
		if !fi.IsDir() {
			fullName := pathName + "/" + fi.Name()
			result = append(result, fullName)
		}
	}
	return result, nil
}

// ReadExcel 读取 excel 并返回所有行数据
func ReadExcel(fileName string) ([][]string, error) {
	var f *excelize.File
	var err error
	var num = 1
	// 每个文件尝试打开10次,10次之后打不开就放弃
	for num <= 10 {
		f, err = excelize.OpenFile(fileName)
		if err != nil {
			time.Sleep(1 * time.Second)
			log.Printf("打开文件错误:{%v} ,睡眠 1s 第{%v}次尝试打开文件{%v}", err, num, fileName)
			num += 1
			continue
		}
		break
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Println(err)
		}
	}()
	if err != nil {
		return nil, err
	}
	firstSheet := f.GetSheetName(0)
	rows, err := f.GetRows(firstSheet)
	if err != nil {
		return nil, err
	}
	log.Printf("读取{%v}内容成功\n", fileName)
	return rows, nil
}

func MergeExcels(allFiles []string) error {
	f := excelize.NewFile()

	sheet := f.NewSheet("Sheet1")
	sheetName := f.GetSheetName(sheet)
	num := 1
	for index, file := range allFiles {
		axis := fmt.Sprintf("A%d", num)
		fileRows, err := ReadExcel(file)
		if err != nil {
			log.Printf("读取文件[%s]内容失败", file)
			return err
		}
		var error2 error
		if index == 0 {
			for _, rows := range fileRows {
				error2 = f.SetSheetRow(sheetName, axis, &rows)
				if err != nil {
					log.Printf("写入文件第{%v}行失败,错误:%v", num, err)
					break
				}
				num += 1
				axis = fmt.Sprintf("A%d", num)
			}
		} else {
			for index2, rows := range fileRows {
				if index2 == 0 {
					continue
				}
				error2 = f.SetSheetRow(sheetName, axis, &rows)
				if err != nil {
					log.Printf("写入文件第{%v}行失败,错误:%v", num, err)
					break
				}
				num += 1
				axis = fmt.Sprintf("A%d", num)
			}
		}
		if error2 != nil {
			return error2
		}
	}
	filename := fmt.Sprintf("./result/合并%v.xlsx", uuid.MustString())
	err := f.SaveAs(filename)
	if err != nil {
		log.Printf("合并excel保存失败")
		return err
	}

	return nil
}
