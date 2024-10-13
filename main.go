package main

import (
	"fmt"
	"log"
	"mkmgo-momworks/imunisasi"
	"net/http"
	"os"

	"gopkg.in/yaml.v2"
)

// Config holds the application configuration, including settings.
type Config struct {
	SasaranImunisasiCfg imunisasi.SasaranImunisasiConfig `yaml:"sasaran_imunisasi_config"` // Configuration specific to SasaranImunisasiService
}

// LoadConfig reads the configuration from a YAML file and returns a Config struct.
func LoadConfig() (*Config, error) {
	file, err := os.Open("config.yaml")
	if err != nil {
		return nil, fmt.Errorf("error opening file: %w", err)
	}
	defer file.Close()

	decoder := yaml.NewDecoder(file)
	var cfg Config
	if err := decoder.Decode(&cfg); err != nil {
		return nil, fmt.Errorf("error decoding YAML: %w", err)
	}

	return &cfg, nil
}

func main() {
	// Load the configuration
	cfg, err := LoadConfig()
	if err != nil {
		log.Fatalf("Failed to load config: %v", err)
	}

	// Initialize the handler with Sasaran Imunisasi services
	sasaranImunisasiHandler := imunisasi.NewSasaranImunisasiHandler(imunisasi.NewSasaranImunisasiService(&cfg.SasaranImunisasiCfg))

	// Define the route and handler for generating files
	http.HandleFunc("/momworks/sasaran/imunisasi", sasaranImunisasiHandler.GenerateFileHandler)

	log.Println("Starting momworks server on localhost:8080...")
	if err := http.ListenAndServe("localhost:8080", nil); err != nil {
		log.Fatalf("Server failed: %v", err)
	}
}
