package routes

import (
	"excel-copy/handlers"

	"github.com/gofiber/fiber/v2"
	"github.com/gofiber/fiber/v2/middleware/logger"
)

func SetupRoutes(app *fiber.App) {
	app.Use(logger.New())

	api := app.Group("/api")

	// POST /api/excel/merge
	api.Post("/excel/merge", handlers.HandleExcelMerge)
}
