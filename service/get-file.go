package service

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// --- ฟังก์ชันช่วยแสดงข้อมูลสำหรับ debug ---
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

// func ShowSourceContent() {
// 	f, err := excelize.OpenFile("source.xlsx")
// 	if err != nil {
// 		fmt.Println("เปิดไฟล์ต้นทางไม่ได้:", err)
// 		return
// 	}
// 	defer f.Close()

// 	rows, err := f.GetRows("Sheet1")
// 	if err != nil {
// 		fmt.Println("อ่านข้อมูลไฟล์ต้นทางไม่ได้:", err)
// 		return
// 	}

// 	fmt.Println("\nเนื้อหาไฟล์ source.xlsx:")
// 	for i, row := range rows {
// 		if i >= 10 { // แสดงแค่ 10 แถวแรก
// 			fmt.Println("...")
// 			break
// 		}
// 		fmt.Printf("แถว %d: %v\n", i+1, row)
// 	}
// }

// func ShowTargetContent() {
// 	f, err := excelize.OpenFile("target.xlsx")
// 	if err != nil {
// 		fmt.Println("เปิดไฟล์ปลายทางไม่ได้:", err)
// 		return
// 	}
// 	defer f.Close()

// 	rows, err := f.GetRows("Result")
// 	if err != nil {
// 		fmt.Println("อ่านข้อมูลไฟล์ปลายทางไม่ได้:", err)
// 		return
// 	}

// 	fmt.Println("\nเนื้อหาไฟล์ target.xlsx:")
// 	for i, row := range rows {
// 		if i >= 10 { // แสดงแค่ 10 แถวแรก
// 			fmt.Println("...")
// 			break
// 		}
// 		fmt.Printf("แถว %d: %v\n", i+1, row)
// 	}
// }
