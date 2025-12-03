package configs

import (
	"fmt"
	"os"
	"time"

	"github.com/joho/godotenv"
)

func Init() {
	initEnvLoader()
	initTimeZone()
}

func initEnvLoader() {
	if err := godotenv.Load(); err != nil {
		fmt.Println("No .env file found, using system environment variables only")
	} else {
		fmt.Println(".env file loaded successfully")
	}
}

func initTimeZone() {
	ict, err := time.LoadLocation("Asia/Bangkok")
	if err != nil {
		fmt.Printf("Error loading timezone: %v\n", err)
		os.Exit(1)
	}
	time.Local = ict
}
