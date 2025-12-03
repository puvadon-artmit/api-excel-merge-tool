package main

import (
	"excel-copy/configs"
	"excel-copy/middlewares"
	"excel-copy/pkgs/logs"
	"excel-copy/routes"
	"fmt"
	"os"

	"github.com/gofiber/fiber/v2"
)

func init() {
	configs.Init()
	logs.LogInit()
}

func main() {
	app := fiber.New()

	app.Use(
		middlewares.NewLoggerMiddleware,
		middlewares.NewCorsMiddleware,
	)

	routes.SetupRoutes(app)

	port := os.Getenv("SERVER_PORT")
	if port == "" {
		port = "8090"
	}
	fmt.Println("Server started at :", port)

	if err := app.Listen(":" + port); err != nil {
		fmt.Printf("Error starting server: %v\n", err)
	}
}
