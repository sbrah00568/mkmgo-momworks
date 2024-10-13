package imunisasi

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

// SasaranImunisasiHandler handles HTTP requests for generating Excel files.
type SasaranImunisasiHandler struct {
	SasaranImunisasiService XlsxFileTransformer
}

// NewSasaranImunisasiHandler initializes a new SasaranImunisasiHandler.
func NewSasaranImunisasiHandler(svc XlsxFileTransformer) *SasaranImunisasiHandler {
	return &SasaranImunisasiHandler{
		SasaranImunisasiService: svc,
	}
}

const (
	maxUploadSize    = 10 << 20 // 10 MB
	fileFormField    = "myFile"
	sheetFormField   = "sheetName"
	sasaranTypeField = "sasaranType"
)

// GenerateFileHandler handles file uploads and generates a new Excel file.
func (h *SasaranImunisasiHandler) GenerateFileHandler(w http.ResponseWriter, r *http.Request) {
	if err := r.ParseMultipartForm(maxUploadSize); err != nil {
		http.Error(w, "File size too large", http.StatusBadRequest)
		log.Printf("File upload error: %v", err)
		return
	}

	// Handle file upload
	tempFilePath, err := HandleFileUpload(r)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	defer os.Remove(tempFilePath)

	// Retrieves the xlsx source file
	ctx := context.WithValue(r.Context(), sasaranTypeKey, r.FormValue(sasaranTypeField))
	sourceFile, err := GetXlsxSourceFile(tempFilePath, r.FormValue(sheetFormField), ctx)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	// Generate the new xlsx file
	generatedFile, err := h.SasaranImunisasiService.GenerateFile(*sourceFile)
	if err != nil {
		http.Error(w, "Error creating file", http.StatusInternalServerError)
		return
	}

	// Set response headers for file download
	if err := WriteXlsxFileToResponse(w, generatedFile); err != nil {
		http.Error(w, "Unable to generate file", http.StatusInternalServerError)
		return
	}
}

// HandleFileUpload manages file upload and returns the file path and error.
func HandleFileUpload(r *http.Request) (string, error) {
	src, fileHeader, err := r.FormFile(fileFormField)
	if err != nil {
		log.Printf("Error retrieving file: %v", err)
		return EMPTY_STRING, fmt.Errorf("failed to retrieve file")
	}
	defer src.Close()

	log.Printf("Uploaded File: %s, Size: %d, MIME: %v", fileHeader.Filename, fileHeader.Size, fileHeader.Header)

	// Create a temporary file
	tempFile, err := os.CreateTemp("temp", filepath.Base(fileHeader.Filename)+"-*.xlsx")
	if err != nil {
		log.Printf("Error creating temp file: %v", err)
		return EMPTY_STRING, fmt.Errorf("failed to create temp file")
	}
	defer tempFile.Close()

	// Write to temp file
	if _, err := io.Copy(tempFile, src); err != nil {
		log.Printf("Error writing to temp file: %v", err)
		return EMPTY_STRING, fmt.Errorf("failed to write to temp file")
	}

	return tempFile.Name(), nil
}

// GetXlsxSourceFile returns source xlsx file from temp
func GetXlsxSourceFile(tempFilePath, sheetName string, ctx context.Context) (*XlsxSourceFile, error) {
	excelFile, err := excelize.OpenFile(tempFilePath)
	if err != nil {
		log.Printf("Error opening xlsx source file: %v", err)
		return nil, fmt.Errorf("error opening xlsx source file")
	}
	defer excelFile.Close()

	return &XlsxSourceFile{
		Ctx:          ctx,
		TempFilePath: tempFilePath,
		SheetName:    sheetName,
		ExcelizeFile: excelFile,
	}, nil
}

// WriteToResponse sets the headers and writes the generated Excel file to the response.
func WriteXlsxFileToResponse(w http.ResponseWriter, generatedFile *XlsxGeneratedFile) error {
	// Set response headers for file download
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", fmt.Sprintf(`attachment; filename="%s"`, generatedFile.FileName))

	// Write the generated Excel file to the response
	if err := generatedFile.ExcelizeFile.Write(w); err != nil {
		log.Printf("Error writing generated file to response: %v", err)
		return fmt.Errorf("failed to write Excel file to response: %w", err)
	}

	log.Printf("Successfully uploaded and processed file: %s", generatedFile.FileName)
	return nil
}
