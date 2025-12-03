package handlers

import (
	"excel-copy/service"
	"fmt"
	"os"
	"path/filepath"

	"github.com/gofiber/fiber/v2"
)

var mappingsResult []service.CellMapping = []service.CellMapping{
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

var sumMappingsResult []service.SumMapping = []service.SumMapping{
	// P6 AVP
	{From: []string{"EX52", "EX55", "EX56", "EX60", "EX61", "EX62", "EX63"}, To: "AG25"},

	// P6 VIP
	{From: []string{"EY52", "EY55", "EY56", "EY60", "EY61", "EY62", "EY63"}, To: "AC25"},

	// P6 TT
	{From: []string{"EZ52", "EZ55", "EZ56", "EZ60", "EZ61", "EZ62", "EZ63"}, To: "AA25"},
}

var mappings2Result []service.CellMapping = []service.CellMapping{
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

var sumMappings2Result []service.SumMapping = []service.SumMapping{
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

func HandleExcelMerge(c *fiber.Ctx) error {
	// รับไฟล์ source
	sourceHeader, err := c.FormFile("source")
	if err != nil {
		c.Locals("error_detail", err.Error())
		return c.Status(fiber.StatusBadRequest).JSON(fiber.Map{
			"error":  "กรุณาอัปโหลดไฟล์ source (field: source)",
			"detail": err.Error(),
		})
	}

	// รับไฟล์ target
	targetHeader, err := c.FormFile("target")
	if err != nil {
		c.Locals("error_detail", err.Error())
		fmt.Println("ERROR: ไม่พบไฟล์ target:", err)
		return c.Status(fiber.StatusBadRequest).JSON(fiber.Map{
			"error":  "กรุณาอัปโหลดไฟล์ target (field: target)",
			"detail": err.Error(),
		})
	}

	// เตรียม path ชั่วคราว
	tmpDir := os.TempDir()
	sourceTmpPath := filepath.Join(tmpDir, "source-"+sourceHeader.Filename)
	targetTmpPath := filepath.Join(tmpDir, "target-"+targetHeader.Filename)

	// บันทึกไฟล์ source
	if err := c.SaveFile(sourceHeader, sourceTmpPath); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "บันทึกไฟล์ source ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	// บันทึกไฟล์ target
	if err := c.SaveFile(targetHeader, targetTmpPath); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "บันทึกไฟล์ target ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	defer os.Remove(sourceTmpPath)
	defer os.Remove(targetTmpPath)

	// เรียก service
	svc := service.ExcelCopyService{
		SourceFile: sourceTmpPath,
		TargetFile: targetTmpPath,
	}

	// =========================== Sheet D070 ====================================

	if err := svc.CopyByCellMapping("Act", "D070", mappingsResult); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (คัดลอก cell ชีต Result) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	if err := svc.SumByCellMapping("Act", "D070", sumMappingsResult); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (รวมยอด ชีต Result) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	// =========================== Sheet D071 ====================================

	if err := svc.CopyByCellMapping("Act", "D071", mappings2Result); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (คัดลอก cell ชีต Result) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	if err := svc.SumByCellMapping("Act", "D071", sumMappings2Result); err != nil {
		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
			"error":  "ประมวลผล (รวมยอด ชีต Result) ไม่สำเร็จ",
			"detail": err.Error(),
		})
	}

	return c.Download(targetTmpPath, "target-merged.xlsx")
}

// func HandleExcelMerge(c *fiber.Ctx) error {
// 	// ---- 1) รับไฟล์จากหน้าบ้าน ----
// 	sourceHeader, err := c.FormFile("source")
// 	if err != nil {
// 		return c.Status(fiber.StatusBadRequest).JSON(fiber.Map{
// 			"error": "กรุณาอัปโหลดไฟล์ source (field: source)",
// 		})
// 	}

// 	targetHeader, err := c.FormFile("target")
// 	if err != nil {
// 		return c.Status(fiber.StatusBadRequest).JSON(fiber.Map{
// 			"error": "กรุณาอัปโหลดไฟล์ target (field: target)",
// 		})
// 	}

// 	// ---- 2) เซฟไฟล์ลง temp ----
// 	tmpDir := os.TempDir()

// 	sourceTmpPath := filepath.Join(tmpDir, "source-"+sourceHeader.Filename)
// 	targetTmpPath := filepath.Join(tmpDir, "target-"+targetHeader.Filename)

// 	if err := c.SaveFile(sourceHeader, sourceTmpPath); err != nil {
// 		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
// 			"error":  "บันทึกไฟล์ source ไม่สำเร็จ",
// 			"detail": err.Error(),
// 		})
// 	}

// 	if err := c.SaveFile(targetHeader, targetTmpPath); err != nil {
// 		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
// 			"error":  "บันทึกไฟล์ target ไม่สำเร็จ",
// 			"detail": err.Error(),
// 		})
// 	}

// 	defer os.Remove(sourceTmpPath)
// 	defer os.Remove(targetTmpPath)

// 	// ---- 3) เรียก service เพื่อ merge ----
// 	svc := service.ExcelCopyService{
// 		SourceFile: sourceTmpPath,
// 		TargetFile: targetTmpPath,
// 	}

// 	colsToCopy := []string{"A", "C", "D"}
// 	targetHeaderArr := []string{"ไอดี", "อายุ_EM", "เมือง_EM", "เพศ", "สีผม"}

// 	if err := svc.CopySelectedColumns(colsToCopy, targetHeaderArr); err != nil {
// 		return c.Status(fiber.StatusInternalServerError).JSON(fiber.Map{
// 			"error":  "ประมวลผลไม่สำเร็จ",
// 			"detail": err.Error(),
// 		})
// 	}

// 	// ---- 4) ส่งไฟล์ target ที่ merge แล้วกลับไปให้ดาวน์โหลด ----
// 	return c.Download(targetTmpPath, "target-merged.xlsx")
// }
