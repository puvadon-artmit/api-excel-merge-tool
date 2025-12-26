package service

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func ShowSourceContent(filePath string) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println("เปิดไฟล์ต้นทางไม่สำเร็จ:", err)
		return
	}
	rows, _ := f.GetRows("Sheet1")
	fmt.Println("เนื้อหาไฟล์ต้นทาง:")
	for _, r := range rows {
		fmt.Println(r)
	}
}

func ShowTargetContent(filePath string) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println("เปิดไฟล์ปลายทางไม่สำเร็จ:", err)
		return
	}
	rows, _ := f.GetRows("Result")
	fmt.Println("เนื้อหาไฟล์ปลายทาง:")
	for _, r := range rows {
		fmt.Println(r)
	}
}
