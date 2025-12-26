package handlers

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/gofiber/fiber/v2"
	"github.com/xuri/excelize/v2"
)

var mappingsResult []CellMapping = []CellMapping{
	// P1
	{From: "EZ16", To: "AA18"},
	{From: "EY16", To: "AC18"},
	{From: "EX16", To: "AG18"},

	// P2
	{From: "EX49", To: "AG19"},
	{From: "EY49", To: "AC19"},
	{From: "EZ49", To: "AA19"},

	// P3
	{From: "EX51", To: "AG22"},
	{From: "EY51", To: "AC22"},
	{From: "EZ51", To: "AA22"},

	// P4
	{From: "EX54", To: "AG23"},
	{From: "EY54", To: "AC23"},
	{From: "EZ54", To: "AA23"},

	// P5
	{From: "EX59", To: "AG24"},
	{From: "EY59", To: "AC24"},
	{From: "EZ59", To: "AA24"},

	// P7
	{From: "EZ67", To: "AA31"},
}

var sumMappingsResult []SumMapping = []SumMapping{
	// P6 AVP
	{From: []string{"EX52", "EX55", "EX56", "EX60", "EX61", "EX62", "EX63"}, To: "AG25"},

	// P6 VIP
	{From: []string{"EY52", "EY55", "EY56", "EY60", "EY61", "EY62", "EY63"}, To: "AC25"},

	// P6 TT
	{From: []string{"EZ52", "EZ55", "EZ56", "EZ60", "EZ61", "EZ62", "EZ63"}, To: "AA25"},
}

var mappings2Result []CellMapping = []CellMapping{
	// P1
	{From: "EY7", To: "AD18"},
	{From: "EX7", To: "AH18"},

	// P2
	{From: "EY20", To: "AD21"},
	{From: "EX20", To: "AH21"},

	// P3
	{From: "EY21", To: "AD23"},
	{From: "EX21", To: "AH23"},

	// P4
	{From: "EY23", To: "AD25"},
	{From: "EX23", To: "AH25"},

	// P6
	{From: "EY32", To: "AD28"},
	{From: "EX32", To: "AH28"},

	// P7
	{From: "EY35", To: "AD29"},
	{From: "EX35", To: "AH29"},

	// P9
	{From: "EY44", To: "AD34"},
	{From: "EX44", To: "AH34"},
}

var sumMappings2Result []SumMapping = []SumMapping{
	// P5
	{From: []string{"EX24", "EX25", "EX26", "EX27", "EX28", "EX29", "EX30"}, To: "AH26"},
	{From: []string{"EY24", "EY25", "EY26", "EY27", "EY28", "EY29", "EY30"}, To: "AD26"},

	// P8
	{From: []string{"EX38", "EX39", "EX40"}, To: "AH30"},
	{From: []string{"EY38", "EY39", "EY40"}, To: "AD30"},

	// P10
	{From: []string{"EX47", "EX48"}, To: "AH36"},
	{From: []string{"EY47", "EY48"}, To: "AD36"},
}

type ExcelCopyService struct {
	SourceFile string
	TargetFile string

	source       *excelize.File
	target       *excelize.File
	targetExists bool
	opened       bool
}

type CellMapping struct {
	From string
	To   string
}

type SumMapping struct {
	From []string
	To   string
}

// ---------- helpers ----------
func ptrBool(v bool) *bool       { return &v }
func ptrString(v string) *string { return &v }

func (s *ExcelCopyService) Open() error {
	if s.opened {
		return nil
	}

	src, err := excelize.OpenFile(s.SourceFile)
	if err != nil {
		return fmt.Errorf("เปิดไฟล์ต้นทางไม่สำเร็จ: %w", err)
	}

	var tgt *excelize.File
	targetExists := false
	if _, err := os.Stat(s.TargetFile); err == nil {
		targetExists = true
		tgt, err = excelize.OpenFile(s.TargetFile)
		if err != nil {
			_ = src.Close()
			return fmt.Errorf("เปิดไฟล์ปลายทางไม่สำเร็จ: %w", err)
		}
	} else {
		tgt = excelize.NewFile()
	}

	s.source = src
	s.target = tgt
	s.targetExists = targetExists
	s.opened = true
	return nil
}

func (s *ExcelCopyService) Close() {
	if s.source != nil {
		_ = s.source.Close()
	}
	if s.target != nil {
		_ = s.target.Close()
	}
}

func (s *ExcelCopyService) ensureSheet(sheetName string) (int, error) {
	idx, err := s.target.GetSheetIndex(sheetName)
	if err == nil && idx != -1 {
		return idx, nil
	}
	idx, err = s.target.NewSheet(sheetName)
	if err != nil {
		return -1, fmt.Errorf("สร้างชีต %s ไม่สำเร็จ: %w", sheetName, err)
	}
	return idx, nil
}

// ✅ Save ครั้งเดียว + บังคับให้ Excel คำนวณสูตรตอนเปิดไฟล์
func (s *ExcelCopyService) Save() error {
	if !s.opened || s.target == nil {
		return fmt.Errorf("ยังไม่ได้ Open() service")
	}

	// ✅ excelize ที่ถูกต้องคือ SetCalcProps / CalcPropsOptions
	if err := s.target.SetCalcProps(&excelize.CalcPropsOptions{
		CalcMode:       ptrString("auto"),
		FullCalcOnLoad: ptrBool(true),
	}); err != nil {
		return fmt.Errorf("ตั้งค่า calc props ไม่สำเร็จ: %w", err)
	}

	if s.targetExists {
		if err := s.target.Save(); err != nil {
			return fmt.Errorf("บันทึกไฟล์ปลายทางล้มเหลว: %w", err)
		}
		return nil
	}

	if err := s.target.SaveAs(s.TargetFile); err != nil {
		return fmt.Errorf("บันทึกไฟล์ปลายทางล้มเหลว: %w", err)
	}
	s.targetExists = true
	return nil
}

// ✔ copy cell (หารพัน + ปัดเศษ)
func (s *ExcelCopyService) CopyByCellMapping(sourceSheet, targetSheet string, mappings []CellMapping) error {
	if err := s.Open(); err != nil {
		return err
	}

	idx, err := s.ensureSheet(targetSheet)
	if err != nil {
		return err
	}
	s.target.SetActiveSheet(idx)

	for _, m := range mappings {
		raw, err := s.source.GetCellValue(sourceSheet, m.From)
		if err != nil {
			return fmt.Errorf("อ่านค่าจาก %s ไม่สำเร็จ: %w", m.From, err)
		}

		raw = strings.TrimSpace(raw)
		if raw == "" {
			if err := s.target.SetCellValue(targetSheet, m.To, ""); err != nil {
				return fmt.Errorf("เขียนค่าที่ %s ไม่สำเร็จ: %w", m.To, err)
			}
			continue
		}

		// ตรวจสอบว่าเป็นค่าลบในรูปแบบวงเล็บ (2,990) หรือไม่
		isNegative := false
		if strings.HasPrefix(raw, "(") && strings.HasSuffix(raw, ")") {
			isNegative = true
			raw = strings.TrimPrefix(raw, "(")
			raw = strings.TrimSuffix(raw, ")")
		}

		numStr := strings.ReplaceAll(raw, ",", "")
		num, err := strconv.ParseFloat(numStr, 64)
		if err != nil {
			if err := s.target.SetCellValue(targetSheet, m.To, raw); err != nil {
				return fmt.Errorf("เขียนค่าที่ %s ไม่สำเร็จ: %w", m.To, err)
			}
			continue
		}

		if isNegative {
			num = -num
		}

		num = math.Round(num / 1000.0)

		if err := s.target.SetCellValue(targetSheet, m.To, num); err != nil {
			return fmt.Errorf("เขียนค่าที่ %s ไม่สำเร็จ: %w", m.To, err)
		}
	}

	return nil
}

// ✔ sum หลาย cell แล้วเขียนลง target (หารพัน + ปัดเศษ)
func (s *ExcelCopyService) SumByCellMapping(sourceSheet, targetSheet string, mappings []SumMapping) error {
	if err := s.Open(); err != nil {
		return err
	}

	idx, err := s.ensureSheet(targetSheet)
	if err != nil {
		return err
	}
	s.target.SetActiveSheet(idx)

	for _, mp := range mappings {
		var sum float64

		for _, fromCell := range mp.From {
			raw, err := s.source.GetCellValue(sourceSheet, fromCell)
			if err != nil {
				return fmt.Errorf("อ่านค่าจาก %s ไม่สำเร็จ: %w", fromCell, err)
			}
			raw = strings.TrimSpace(raw)
			if raw == "" {
				continue
			}

			// ตรวจสอบว่าเป็นค่าลบในรูปแบบวงเล็บ (2,990) หรือไม่
			isNegative := false
			if strings.HasPrefix(raw, "(") && strings.HasSuffix(raw, ")") {
				isNegative = true
				raw = strings.TrimPrefix(raw, "(")
				raw = strings.TrimSuffix(raw, ")")
			}

			numStr := strings.ReplaceAll(raw, ",", "")
			num, err := strconv.ParseFloat(numStr, 64)
			if err != nil {
				// ถ้าเจอข้อความ ให้ข้าม (หรือจะ return error ก็ได้)
				continue
			}

			if isNegative {
				num = -num
			}

			sum += num
		}

		sum = sum / 1000

		if err := s.target.SetCellValue(targetSheet, mp.To, sum); err != nil {
			return fmt.Errorf("เขียนค่าที่ %s ไม่สำเร็จ: %w", mp.To, err)
		}
	}

	return nil
}

func HandleExcelMerge(c *fiber.Ctx) error {
	sourceHeader, err := c.FormFile("source")
	if err != nil {
		c.Locals("error_detail", err.Error())
		return c.Status(fiber.StatusBadRequest).JSON(fiber.Map{
			"error":  "กรุณาอัปโหลดไฟล์ source (field: source)",
			"detail": err.Error(),
		})
	}

	targetHeader, err := c.FormFile("target")
	if err != nil {
		c.Locals("error_detail", err.Error())
		return c.Status(fiber.StatusBadRequest).JSON(fiber.Map{
			"error":  "กรุณาอัปโหลดไฟล์ target (field: target)",
			"detail": err.Error(),
		})
	}

	tmpDir := os.TempDir()
	sourceTmpPath := filepath.Join(tmpDir, "source-"+sourceHeader.Filename)
	targetTmpPath := filepath.Join(tmpDir, "target-"+targetHeader.Filename)

	if err := c.SaveFile(sourceHeader, sourceTmpPath); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "บันทึกไฟล์ source ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	if err := c.SaveFile(targetHeader, targetTmpPath); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "บันทึกไฟล์ target ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	defer os.Remove(sourceTmpPath)
	defer os.Remove(targetTmpPath)

	svc := &ExcelCopyService{
		SourceFile: sourceTmpPath,
		TargetFile: targetTmpPath,
	}

	// เปิดครั้งเดียว
	if err := svc.Open(); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "เปิดไฟล์ Excel ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}
	defer svc.Close()

	// D070
	if err := svc.CopyByCellMapping("Act", "D070", mappingsResult); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (คัดลอก cell ชีต D070) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}
	if err := svc.SumByCellMapping("Act", "D070", sumMappingsResult); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (รวมยอด ชีต D070) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	// D071
	if err := svc.CopyByCellMapping("Act", "D071", mappings2Result); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (คัดลอก cell ชีต D071) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}
	if err := svc.SumByCellMapping("Act", "D071", sumMappings2Result); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (รวมยอด ชีต D071) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	// ✅ Save ทีเดียว (และใน Save มี SetCalcProps แล้ว)
	if err := svc.Save(); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "บันทึกไฟล์ target ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	return c.Download(targetTmpPath, "target-merged.xlsx")
}
