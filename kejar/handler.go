package kejar

import (
	"context"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

// SasaranImunisasiHandler handles the HTTP requests for generating Excel files,
// utilizing services that implement XlsxFileTransformer interface for processing bayi and baduta data.
type SasaranImunisasiHandler struct {
	SasaranImunisasiService XlsxFileTransformer
}

// NewSasaranImunisasiHandler initializes a new SasaranImunisasiHandler with the provided services.
func NewSasaranImunisasiHandler(sasaranImunisasiSvc XlsxFileTransformer) *SasaranImunisasiHandler {
	return &SasaranImunisasiHandler{
		SasaranImunisasiService: sasaranImunisasiSvc,
	}
}

// GenerateFileHandler handles file uploads and generates the new Sasaran Imunisasi Excel file based on the input.
// The input is data imunisasi anak which can be either bayi or baduta. Based on thee data, this api will filter and validate
// to retrieve Sasaran Imunisasi which contains data anak that yet to receive complete immunization.
func (h *SasaranImunisasiHandler) GenerateFileHandler(w http.ResponseWriter, r *http.Request) {

	// Parse the multipart form, checking for size constraints.
	const maxUploadSize = 10 << 20 // 10 MB
	if err := r.ParseMultipartForm(maxUploadSize); err != nil {
		http.Error(w, "File size too large", http.StatusBadRequest)
		log.Printf("File upload error: %v", err)
		return
	}

	// Retrieve the uploaded file from the form.
	src, fileHeader, err := r.FormFile("myFile")
	if err != nil {
		http.Error(w, "Failed to retrieve file", http.StatusBadRequest)
		log.Printf("Error retrieving file: %v", err)
		return
	}
	defer src.Close()

	log.Printf("Uploaded File: %s, Size: %d, MIME: %v", fileHeader.Filename, fileHeader.Size, fileHeader.Header)

	// Create temporary location to save the uploaded file.
	tempFile, err := os.CreateTemp("temp", filepath.Base(fileHeader.Filename)+"-*.xlsx")
	if err != nil {
		http.Error(w, "Failed to create temp file", http.StatusInternalServerError)
		log.Printf("Error to create temp file: %v", err)
		return
	}
	defer tempFile.Close()

	// Write uploaded file to temporary location (Save).
	if _, err := io.Copy(tempFile, src); err != nil {
		http.Error(w, "Failed to write to temp file file", http.StatusInternalServerError)
		log.Printf("Error writing to temp file file: %v", err)
		return
	}

	tempFilePath := tempFile.Name()
	defer os.Remove(tempFilePath)

	// Open the Excel file.
	excelizeFile, err := excelize.OpenFile(tempFilePath)
	if err != nil {
		http.Error(w, "Failed to open file", http.StatusInternalServerError)
		log.Printf("Error opening Excel file: %v", err)
		return
	}
	defer excelizeFile.Close()

	sourceFile := SourceXlsxFile{
		Ctx:          context.WithValue(r.Context(), sasaranTypeKey, r.FormValue("sasaranType")),
		TempFilePath: tempFilePath,
		SheetName:    r.FormValue("sheetName"),
		ExcelizeFile: excelizeFile,
	}
	generatedFile, err := h.SasaranImunisasiService.GenerateFile(sourceFile)
	if err != nil {
		http.Error(w, "Error creating file", http.StatusInternalServerError)
		log.Printf("Error generating file: %v", err)
		return
	}

	// Set the response headers for file download.
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", fmt.Sprintf(`attachment; filename="%s"`, generatedFile.FileName))

	// Write the generated Excel file to the response.
	if err := generatedFile.ExcelizeFile.Write(w); err != nil {
		http.Error(w, "Unable to generate file", http.StatusInternalServerError)
		return
	}

	log.Println("Successfully uploaded and processed file")
}
