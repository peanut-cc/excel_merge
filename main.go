package main

import (
	"log"
	"os"
	"path/filepath"
)

func init() {
	os.MkdirAll("./result", os.ModePerm)
	os.MkdirAll("./src_files", os.ModePerm)
}

func main() {
	srcfilesAbsPath, err := filepath.Abs("./src_files")
	if err != nil {
		log.Printf("get src_fiels dir error:%v", err)
		return
	}
	allFiles, err := GetAllFiles(srcfilesAbsPath)
	if err != nil {
		return
	}
	if len(allFiles) == 0 {
		log.Fatalf("路径[%s]下没有文件", srcfilesAbsPath)
	}
	err = MergeExcels(allFiles)
	if err != nil {
		return
	}
	log.Println("success!!!")
}
