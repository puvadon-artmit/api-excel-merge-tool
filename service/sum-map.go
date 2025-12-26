package service

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func parseExcelFloat(raw string) (float64, error) {
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return 0, nil
	}

	isNegative := false

	if strings.HasPrefix(raw, "(") && strings.HasSuffix(raw, ")") {
		isNegative = true
		raw = strings.TrimPrefix(raw, "(")
		raw = strings.TrimSuffix(raw, ")")
		raw = strings.TrimSpace(raw)
	}

	// ตัด comma ออก เผื่อเป็น 12,345.67
	raw = strings.ReplaceAll(raw, ",", "")

	v, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		return 0, err
	}

	if isNegative {
		v = -v
	}

	return v, nil
}

type SumMapping struct {
	From []string // เซลล์ต้นทางหลายเซลล์ เช่น {"B27","B28","B29"}
	To   string   // เซลล์ปลายทาง เช่น "C30"
}

// SumByCellMapping: รวมค่าตัวเลขจากหลายเซลล์แล้วใส่ลงเซลล์ปลายทาง (หารพัน + ปัดเศษ)
func (s *ExcelCopyService) SumByCellMapping(
	sourceSheet string,
	targetSheet string,
	mappings []SumMapping,
) error {
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

	// 3) loop ตาม mappings → รวมยอดแล้วเขียนลงเซลล์ปลายทาง
	for _, m := range mappings {
		var sum float64

		for _, fromCell := range m.From {
			raw, err := source.GetCellValue(sourceSheet, fromCell)
			if err != nil {
				return fmt.Errorf("อ่านค่าจาก %s ไม่สำเร็จ: %w", fromCell, err)
			}

			raw = strings.TrimSpace(raw)
			if raw == "" {
				continue
			}

			v, err := parseExcelFloat(raw)
			if err != nil {
				return fmt.Errorf("แปลงค่าที่ %s (%s) เป็นตัวเลขไม่ได้: %w", fromCell, raw, err)
			}
			sum += v
		}

		// ----- หารพัน + ปัดเศษ (.5 ขึ้น) -----
		// sum = sum / 1000.0
		// sum = math.Floor(sum + 0.1)

		if err := target.SetCellValue(targetSheet, m.To, sum); err != nil {
			return fmt.Errorf("เขียนผลรวมที่ %s ไม่สำเร็จ: %w", m.To, err)
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
