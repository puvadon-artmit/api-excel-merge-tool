package service

import (
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelCopyService struct {
	SourceFile string
	TargetFile string
}

type CellMapping struct {
	From string // เซลล์ต้นทาง เช่น "AA31"
	To   string // เซลล์ปลายทาง เช่น "G10" หรือ "AA31"
}

// ✔ ก็อปแบบ 1:1 เซลล์เดียว (หารพัน + ปัดเศษ)
func (s *ExcelCopyService) CopyByCellMapping(
	sourceSheet string,
	targetSheet string,
	mappings []CellMapping,
) error {

	fmt.Println("🚀 CopyByCellMapping", mappings)
	// 1) เปิดไฟล์ source
	source, err := excelize.OpenFile(s.SourceFile)
	if err != nil {
		return fmt.Errorf("เปิดไฟล์ต้นทางไม่สำเร็จ: %w", err)
	}
	defer source.Close()

	// 2) เปิดหรือสร้างไฟล์ target
	var target *excelize.File
	targetExists := false

	if _, err := os.Stat(s.TargetFile); err == nil {
		targetExists = true
		target, err = excelize.OpenFile(s.TargetFile)
		if err != nil {
			return fmt.Errorf("เปิดไฟล์ปลายทางไม่สำเร็จ: %w", err)
		}
	} else {
		target = excelize.NewFile()
	}
	defer target.Close()

	// ถ้าไม่มีชีต targetSheet ให้สร้าง
	idx, err := target.GetSheetIndex(targetSheet)
	if err != nil || idx == -1 {
		idx, err = target.NewSheet(targetSheet)
		if err != nil {
			return fmt.Errorf("สร้างชีต %s ไม่สำเร็จ: %w", targetSheet, err)
		}
	}
	target.SetActiveSheet(idx)

	// 3) loop ตาม mappings แล้วก็อปทีละเซลล์ (หารพัน + ปัดเศษ)
	for _, m := range mappings {
		raw, err := source.GetCellValue(sourceSheet, m.From)
		if err != nil {
			return fmt.Errorf("อ่านค่าจาก %s ไม่สำเร็จ: %w", m.From, err)
		}

		raw = strings.TrimSpace(raw)
		if raw == "" {
			// ถ้าเป็นค่าว่าง ใส่ว่างกลับไปเลย
			if err := target.SetCellValue(targetSheet, m.To, ""); err != nil {
				return fmt.Errorf("เขียนค่าที่ %s ไม่สำเร็จ: %w", m.To, err)
			}
			continue
		}

		// ลองแปลงเป็นตัวเลข: ตัด comma ออกก่อน เผื่อมี format 12,345.67
		numStr := strings.ReplaceAll(raw, ",", "")
		num, err := strconv.ParseFloat(numStr, 64)
		if err != nil {
			// ถ้าแปลงไม่ได้ ให้ก็อป string เดิม (กันพัง)
			if err := target.SetCellValue(targetSheet, m.To, raw); err != nil {
				return fmt.Errorf("เขียนค่าที่ %s ไม่สำเร็จ: %w", m.To, err)
			}
			continue
		}

		// หารพัน
		num = num / 1000.0

		num = math.Round(num)

		// เขียนเป็นตัวเลขลง target
		if err := target.SetCellValue(targetSheet, m.To, num); err != nil {
			return fmt.Errorf("เขียนค่าที่ %s ไม่สำเร็จ: %w", m.To, err)
		}
	}

	// 4) เซฟไฟล์ปลายทาง
	if targetExists {
		if err := target.Save(); err != nil {
			return fmt.Errorf("บันทึกไฟล์ปลายทางล้มเหลว: %w", err)
		}
	} else {
		if err := target.SaveAs(s.TargetFile); err != nil {
			return fmt.Errorf("บันทึกไฟล์ปลายทางล้มเหลว: %w", err)
		}
	}

	return nil
}
