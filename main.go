package main

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	http.HandleFunc("/momworks/imunisasi/kejar/bayi", GetSasaranImunisasiKejarBayiHandler)
	log.Println("Starting momworks server....")
	if err := http.ListenAndServe("localhost:8080", nil); err != nil {
		log.Fatalf("Server failed: %v", err)
	}
}

// GetSasaranImunisasiKejarBayiHandler processes uploaded Excel files for immunization data.
// It retrieves data based on the provided sheet name and generates a downloadable Excel file.
func GetSasaranImunisasiKejarBayiHandler(w http.ResponseWriter, r *http.Request) {
	const maxUploadSize = 10 << 20 // Set max upload size to 10 MB

	// Parse the multipart form
	if err := r.ParseMultipartForm(maxUploadSize); err != nil {
		http.Error(w, "File size too large", http.StatusBadRequest)
		log.Printf("File upload error: %v", err)
		return
	}

	// Retrieve the file from the form
	file, fileHeader, err := r.FormFile("myFile")
	if err != nil {
		http.Error(w, "Failed to retrieve file", http.StatusBadRequest)
		log.Printf("Error retrieving file: %v", err)
		return
	}
	defer file.Close()

	log.Printf("Uploaded File: %s, Size: %d, MIME: %v", fileHeader.Filename, fileHeader.Size, fileHeader.Header)

	// Save the uploaded file to a temporary location
	tempFilePath, err := saveFileToTemp(file, fileHeader.Filename)
	if err != nil {
		http.Error(w, "Failed to save file", http.StatusInternalServerError)
		log.Printf("Error saving file: %v", err)
		return
	}
	defer os.Remove(tempFilePath) // Clean up temp file after processing

	// Process the Excel file and retrieve data
	dataImunisasiBayi, err := dataImunisasiBayiRetriever(tempFilePath, r.FormValue("sheetName"))
	if err != nil {
		http.Error(w, "Error processing file", http.StatusInternalServerError)
		log.Printf("Error processing Excel file: %v", err)
		return
	}

	// Write data kejar asik bayi excel file
	sasaranImunisasiKejarBayiFile, err := createNewSasaranImunisasiKejarBayiFile(dataImunisasiBayi)
	if err != nil {
		http.Error(w, "Error creating new data kejar asik bayi file", http.StatusInternalServerError)
		log.Printf("Error creating new data kejar asik bayi file: %v", err)
		return
	}

	// Set the content type and disposition for download
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", `attachment; filename="`+getSasaranImunisasiKejarBayiFileName()+`"`)

	// Write the file to the response
	if err := sasaranImunisasiKejarBayiFile.Write(w); err != nil {
		http.Error(w, "Unable to generate file", http.StatusInternalServerError)
		return
	}

	log.Println("Successfully uploaded and processed file")
}

// getSasaranImunisasiKejarBayiFileName returns file name for Sasaran Imunisasi Kejar Bayi
func getSasaranImunisasiKejarBayiFileName() string {
	currentDate := time.Now()

	months := map[time.Month]string{
		time.January:   "Januari",
		time.February:  "Februari",
		time.March:     "Maret",
		time.April:     "April",
		time.May:       "Mei",
		time.June:      "Juni",
		time.July:      "Juli",
		time.August:    "Agustus",
		time.September: "September",
		time.October:   "Oktober",
		time.November:  "November",
		time.December:  "Desember",
	}

	month := months[currentDate.Month()]
	return "Sasaran Imunisasi Kejar Bayi " + strconv.Itoa(currentDate.Day()) + " " + month + ".xlsx"
}

// saveFileToTemp saves the uploaded file to a temporary location and returns the path.
func saveFileToTemp(file io.Reader, filename string) (string, error) {
	tempFile, err := os.CreateTemp("temp", filepath.Base(filename)+"-*.xlsx")
	if err != nil {
		log.Printf("Error creating temp file: %v", err)
		return "", fmt.Errorf("failed to create temp file: %w", err)
	}
	defer tempFile.Close()

	// Copy the file content to the temporary file
	if _, err := io.Copy(tempFile, file); err != nil {
		log.Printf("Error writing to temp file: %v", err)
		return "", fmt.Errorf("failed to write to temp file: %w", err)
	}

	return tempFile.Name(), nil
}

// DataImunisasiBayi represents the vaccination record for bayi.
type DataImunisasiBayi struct {
	ID               string    `json:"id"`
	NikAnak          string    `json:"nikAnak"`
	NamaAnak         string    `json:"namaAnak"`
	TanggalLahirAnak string    `json:"tanggalLahirAnak"`
	JenisKelaminAnak string    `json:"jenisKelaminAnak"`
	NikOrangTua      string    `json:"nikOrangTua"`
	NamaOrangTua     string    `json:"namaOrangTua"`
	Alamat           Alamat    `json:"alamat"`
	Imunisasi        Imunisasi `json:"imunisasi"`
}

// Alamat represents the address details of the family.
type Alamat struct {
	Provinsi      string `json:"provinsi"`
	KabupatenKota string `json:"kabupatenKota"`
	Kecamatan     string `json:"kecamatan"`
	KelurahanDesa string `json:"kelurahanDesa"`
	KodePuskesmas string `json:"kodePuskesmas"`
	Puskesmas     string `json:"puskesmas"`
}

// Imunisasi represents the immunization records for bayi.
type Imunisasi struct {
	HB0       VaksinInfo `json:"hb0"`
	BCG1      VaksinInfo `json:"bcg1"`
	Polio1    VaksinInfo `json:"polio1"`
	Polio2    VaksinInfo `json:"polio2"`
	Polio3    VaksinInfo `json:"polio3"`
	Polio4    VaksinInfo `json:"polio4"`
	DPTHbHib1 VaksinInfo `json:"dpthhb1"`
	DPTHbHib2 VaksinInfo `json:"dpthhb2"`
	DPTHbHib3 VaksinInfo `json:"dpthhb3"`
	IPV1      VaksinInfo `json:"ipv1"`
	IPV2      VaksinInfo `json:"ipv2"`
	ROTA1     VaksinInfo `json:"rota1"`
	ROTA2     VaksinInfo `json:"rota2"`
	ROTA3     VaksinInfo `json:"rota3"`
	PCV1      VaksinInfo `json:"pcv1"`
	PCV2      VaksinInfo `json:"pcv2"`
	JE1       VaksinInfo `json:"je1"`
	MR1       VaksinInfo `json:"mr1"`
	IDL1      VaksinInfo `json:"idl1"`
}

// VaksinInfo contains details about each vaccination.
type VaksinInfo struct {
	Tanggal             string `json:"tanggal"`
	TanggalInput        string `json:"tanggalInput"`
	PosImunisasi        string `json:"posImunisasi"`
	PkmPemberiImunisasi string `json:"pkmPemberiImunisasi"`
	StatusImunisasi     string `json:"statusImunisasi"`
	SumberPencatatan    string `json:"sumberPencatatan"`
}

// dataImunisasiBayiRetriever opens the Excel file, reads the specified sheet, and returns list of data imunisasi bayi.
func dataImunisasiBayiRetriever(filePath, sheetName string) ([]DataImunisasiBayi, error) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Printf("Error opening Excel file: %v", err)
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	defer file.Close()

	// Retrieve rows from the specified sheet
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Printf("Error getting rows from Excel sheet: %v", err)
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}

	return mapRowsToDataImunisasiBayiList(rows, file, sheetName), nil
}

// columnName generates the Excel-style column names for the first n columns.
func columnName(n int) []string {
	var result []string
	for i := 1; i <= n; i++ {
		result = append(result, getColumnName(i))
	}
	return result
}

// getColumnName converts a column number to its Excel-style letter representation.
func getColumnName(i int) string {
	columns := ""
	for i > 0 {
		i--
		columns = string('A'+i%26) + columns
		i /= 26
	}
	return columns
}

// mapRowsToDataImunisasiBayiList converts the rows from the Excel file to a list of DataImunisasiBayi.
func mapRowsToDataImunisasiBayiList(rows [][]string, file *excelize.File, sheetName string) []DataImunisasiBayi {
	var dataImunisasiBayiList []DataImunisasiBayi
	for i := 0; i < len(rows); i++ {
		if i != 0 { // skip header
			adjustedCellValues := cellValuesAdjuster(rows, file, sheetName, i)
			dataImunisasiBayi := mapCellValuesToDataImunisasiBayi(adjustedCellValues)
			if dataImunisasiBayi != nil {
				dataImunisasiBayiList = append(dataImunisasiBayiList, *dataImunisasiBayi)
			}
		}
	}
	return dataImunisasiBayiList
}

// cellValuesAdjuster adjusts cell values when error or blank
func cellValuesAdjuster(rows [][]string, file *excelize.File, sheetName string, rowIndex int) []string {
	var adjustedCellValues []string
	for _, columnNameList := range columnName(len(rows[0])) {
		cell := columnNameList + strconv.Itoa(rowIndex+1)
		cellValue, err := file.GetCellValue(sheetName, cell)
		if err != nil {
			log.Printf("Error getting cell value: %v", err)
			cellValue = "-"
		}

		if cellValue == "" {
			cellValue = "-"
		}
		adjustedCellValues = append(adjustedCellValues, cellValue)
	}
	return adjustedCellValues
}

// mapCellValuesToDataImunisasiBayi converts a single row of strings into a DataImunisasiBayi record.
func mapCellValuesToDataImunisasiBayi(cell []string) *DataImunisasiBayi {
	record := DataImunisasiBayi{
		ID:               cell[0],
		NikAnak:          cell[1],
		NamaAnak:         cell[2],
		TanggalLahirAnak: cell[3],
		JenisKelaminAnak: cell[4],
		NikOrangTua:      cell[5],
		NamaOrangTua:     cell[6],
		Alamat: Alamat{
			Provinsi:      cell[7],
			KabupatenKota: cell[8],
			Kecamatan:     cell[9],
			KelurahanDesa: cell[10],
			KodePuskesmas: cell[11],
			Puskesmas:     cell[12],
		},
		Imunisasi: Imunisasi{
			HB0: VaksinInfo{
				Tanggal:             cell[13],
				TanggalInput:        cell[14],
				PosImunisasi:        cell[15],
				PkmPemberiImunisasi: cell[16],
				StatusImunisasi:     cell[17],
				SumberPencatatan:    cell[18],
			},
			BCG1: VaksinInfo{
				Tanggal:             cell[19],
				TanggalInput:        cell[20],
				PosImunisasi:        cell[21],
				PkmPemberiImunisasi: cell[22],
				StatusImunisasi:     cell[23],
				SumberPencatatan:    cell[24],
			},
			Polio1: VaksinInfo{
				Tanggal:             cell[25],
				TanggalInput:        cell[26],
				PosImunisasi:        cell[27],
				PkmPemberiImunisasi: cell[28],
				StatusImunisasi:     cell[29],
				SumberPencatatan:    cell[30],
			},
			Polio2: VaksinInfo{
				Tanggal:             cell[31],
				TanggalInput:        cell[32],
				PosImunisasi:        cell[33],
				PkmPemberiImunisasi: cell[34],
				StatusImunisasi:     cell[35],
				SumberPencatatan:    cell[36],
			},
			Polio3: VaksinInfo{
				Tanggal:             cell[37],
				TanggalInput:        cell[38],
				PosImunisasi:        cell[39],
				PkmPemberiImunisasi: cell[40],
				StatusImunisasi:     cell[41],
				SumberPencatatan:    cell[42],
			},
			Polio4: VaksinInfo{
				Tanggal:             cell[43],
				TanggalInput:        cell[44],
				PosImunisasi:        cell[45],
				PkmPemberiImunisasi: cell[46],
				StatusImunisasi:     cell[47],
				SumberPencatatan:    cell[48],
			},
			DPTHbHib1: VaksinInfo{
				Tanggal:             cell[49],
				TanggalInput:        cell[50],
				PosImunisasi:        cell[51],
				PkmPemberiImunisasi: cell[52],
				StatusImunisasi:     cell[53],
				SumberPencatatan:    cell[54],
			},
			DPTHbHib2: VaksinInfo{
				Tanggal:             cell[55],
				TanggalInput:        cell[56],
				PosImunisasi:        cell[57],
				PkmPemberiImunisasi: cell[58],
				StatusImunisasi:     cell[59],
				SumberPencatatan:    cell[60],
			},
			DPTHbHib3: VaksinInfo{
				Tanggal:             cell[61],
				TanggalInput:        cell[62],
				PosImunisasi:        cell[63],
				PkmPemberiImunisasi: cell[64],
				StatusImunisasi:     cell[65],
				SumberPencatatan:    cell[66],
			},
			IPV1: VaksinInfo{
				Tanggal:             cell[67],
				TanggalInput:        cell[68],
				PosImunisasi:        cell[69],
				PkmPemberiImunisasi: cell[70],
				StatusImunisasi:     cell[71],
				SumberPencatatan:    cell[72],
			},
			IPV2: VaksinInfo{
				Tanggal:             cell[73],
				TanggalInput:        cell[74],
				PosImunisasi:        cell[75],
				PkmPemberiImunisasi: cell[76],
				StatusImunisasi:     cell[77],
				SumberPencatatan:    cell[78],
			},
			ROTA1: VaksinInfo{
				Tanggal:             cell[79],
				TanggalInput:        cell[80],
				PosImunisasi:        cell[81],
				PkmPemberiImunisasi: cell[82],
				StatusImunisasi:     cell[83],
				SumberPencatatan:    cell[84],
			},
			ROTA2: VaksinInfo{
				Tanggal:             cell[85],
				TanggalInput:        cell[86],
				PosImunisasi:        cell[87],
				PkmPemberiImunisasi: cell[88],
				StatusImunisasi:     cell[89],
				SumberPencatatan:    cell[90],
			},
			ROTA3: VaksinInfo{
				Tanggal:             cell[91],
				TanggalInput:        cell[92],
				PosImunisasi:        cell[93],
				PkmPemberiImunisasi: cell[94],
				StatusImunisasi:     cell[95],
				SumberPencatatan:    cell[96],
			},
			PCV1: VaksinInfo{
				Tanggal:             cell[97],
				TanggalInput:        cell[98],
				PosImunisasi:        cell[99],
				PkmPemberiImunisasi: cell[100],
				StatusImunisasi:     cell[101],
				SumberPencatatan:    cell[102],
			},
			PCV2: VaksinInfo{
				Tanggal:             cell[103],
				TanggalInput:        cell[104],
				PosImunisasi:        cell[105],
				PkmPemberiImunisasi: cell[106],
				StatusImunisasi:     cell[107],
				SumberPencatatan:    cell[108],
			},
			JE1: VaksinInfo{
				Tanggal:             cell[109],
				TanggalInput:        cell[110],
				PosImunisasi:        cell[111],
				PkmPemberiImunisasi: cell[112],
				StatusImunisasi:     cell[113],
				SumberPencatatan:    cell[114],
			},
			MR1: VaksinInfo{
				Tanggal:             cell[115],
				TanggalInput:        cell[116],
				PosImunisasi:        cell[117],
				PkmPemberiImunisasi: cell[118],
				StatusImunisasi:     cell[119],
				SumberPencatatan:    cell[120],
			},
			IDL1: VaksinInfo{
				Tanggal:             cell[121],
				TanggalInput:        cell[122],
				PosImunisasi:        cell[123],
				PkmPemberiImunisasi: cell[124],
				StatusImunisasi:     cell[125],
				SumberPencatatan:    cell[126],
			},
		},
	}

	return &record
}

// SasaranImunisasiKejarBayi represents represents the data of a baby that requires follow-up
// immunization due to missed scheduled vaccinations.
type SasaranImunisasiKejarBayi struct {
	NamaAnak         string `json:"namaAnak"`
	UsiaAnak         string `json:"usiaAnak"`
	TanggalLahirAnak string `json:"tanggalLahirAnak"`
	JenisKelaminAnak string `json:"jenisKelaminAnak"`
	NamaOrangTua     string `json:"namaOrangTua"`
	HBO              int    `json:"hbo"`
	BCG1             int    `json:"bcg1"`
	POLIO1           int    `json:"polio1"`
	POLIO2           int    `json:"polio2"`
	POLIO3           int    `json:"polio3"`
	POLIO4           int    `json:"polio4"`
	DPTHbHib1        int    `json:"dptHbHib1"`
	DPTHbHib2        int    `json:"dptHbHib2"`
	DPTHbHib3        int    `json:"dptHbHib3"`
	IPV1             int    `json:"ipv1"`
	IPV2             int    `json:"ipv2"`
	ROTA1            int    `json:"rota1"`
	ROTA2            int    `json:"rota2"`
	ROTA3            int    `json:"rota3"`
	PCV1             int    `json:"pcv1"`
	PCV2             int    `json:"pcv2"`
	JE1              int    `json:"je1"`
	MR1              int    `json:"mr1"`
	IDL1             int    `json:"idl1"`
}

// newSasaranImunisasiKejarBayiList converts a list of DataImunisasiBayi records into a list of SasaranImunisasiKejarBayi records.
// Only records that pass the validation checks will be added to the resulting list.
func newSasaranImunisasiKejarBayiList(dataImunisasiBayiList []DataImunisasiBayi) []SasaranImunisasiKejarBayi {
	sasaranImunisasiKejarBayiList := make([]SasaranImunisasiKejarBayi, 0, len(dataImunisasiBayiList)-1)
	for i := 1; i < len(dataImunisasiBayiList); i++ { // Skip header
		sasaranImunisasiKejarBayi := newSasaranImunisasiKejarBayi(&dataImunisasiBayiList[i])
		if sasaranImunisasiKejarBayi != nil {
			sasaranImunisasiKejarBayiList = append(sasaranImunisasiKejarBayiList, *sasaranImunisasiKejarBayi)
		}
	}
	return sasaranImunisasiKejarBayiList
}

// newSasaranImunisasiKejarBayi creates a new SasaranImunisasiKejarBayi record from a DataImunisasiBayi object.
// The function calculates the child's age, checks the immunization status, and ensures the data is valid.
func newSasaranImunisasiKejarBayi(dataImunisasiBayi *DataImunisasiBayi) *SasaranImunisasiKejarBayi {
	usiaAnak, err := calculateUsiaAnak(dataImunisasiBayi.TanggalLahirAnak)
	if err != nil {
		return nil
	}

	imunisasi := dataImunisasiBayi.Imunisasi
	if isAllImunisasiIdeal(imunisasi) || isPemberiImunivasiInvalid(imunisasi) || isPosImunivasiInvalid(imunisasi) {
		return nil
	}

	return &SasaranImunisasiKejarBayi{
		NamaAnak:         dataImunisasiBayi.NamaAnak,
		UsiaAnak:         usiaAnak,
		TanggalLahirAnak: dataImunisasiBayi.TanggalLahirAnak,
		JenisKelaminAnak: dataImunisasiBayi.JenisKelaminAnak,
		NamaOrangTua:     dataImunisasiBayi.NamaOrangTua,
		HBO:              getStatusImunisasiKejar(imunisasi.HB0.StatusImunisasi),
		BCG1:             getStatusImunisasiKejar(imunisasi.BCG1.StatusImunisasi),
		POLIO1:           getStatusImunisasiKejar(imunisasi.Polio1.StatusImunisasi),
		POLIO2:           getStatusImunisasiKejar(imunisasi.Polio2.StatusImunisasi),
		POLIO3:           getStatusImunisasiKejar(imunisasi.Polio3.StatusImunisasi),
		POLIO4:           getStatusImunisasiKejar(imunisasi.Polio4.StatusImunisasi),
		DPTHbHib1:        getStatusImunisasiKejar(imunisasi.DPTHbHib1.StatusImunisasi),
		DPTHbHib2:        getStatusImunisasiKejar(imunisasi.DPTHbHib2.StatusImunisasi),
		DPTHbHib3:        getStatusImunisasiKejar(imunisasi.DPTHbHib3.StatusImunisasi),
		IPV1:             getStatusImunisasiKejar(imunisasi.IPV1.StatusImunisasi),
		IPV2:             getStatusImunisasiKejar(imunisasi.IPV2.StatusImunisasi),
		ROTA1:            getStatusImunisasiKejar(imunisasi.ROTA1.StatusImunisasi),
		ROTA2:            getStatusImunisasiKejar(imunisasi.ROTA2.StatusImunisasi),
		ROTA3:            getStatusImunisasiKejar(imunisasi.ROTA3.StatusImunisasi),
		PCV1:             getStatusImunisasiKejar(imunisasi.PCV1.StatusImunisasi),
		PCV2:             getStatusImunisasiKejar(imunisasi.PCV2.StatusImunisasi),
		JE1:              getStatusImunisasiKejar(imunisasi.JE1.StatusImunisasi),
		MR1:              getStatusImunisasiKejar(imunisasi.MR1.StatusImunisasi),
		IDL1:             getStatusImunisasiKejar(imunisasi.IDL1.StatusImunisasi),
	}
}

// calculateUsiaAnak calculates the age of the child based on their birthdate and the current date.
func calculateUsiaAnak(tanggalLahirAnak string) (string, error) {
	birthDate, err := time.Parse("2006-01-02", tanggalLahirAnak)
	if err != nil {
		log.Printf("Failed to parse tanggal lahir anak: %v", err)
		return "", err
	}

	currentDate := time.Now()
	months := currentDate.Year()*12 + int(currentDate.Month()) - (birthDate.Year()*12 + int(birthDate.Month()))
	days := currentDate.Day() - birthDate.Day()

	if days < 0 {
		months--
		days += time.Date(currentDate.Year(), currentDate.Month(), 0, 0, 0, 0, 0, currentDate.Location()).Day()
	}

	return fmt.Sprintf("%d Bulan %d Hari", months, days), nil
}

// getStatusImunisasiKejar returns the immunization status as "0" (ideal) or "1" (need catch-up).
func getStatusImunisasiKejar(status string) int {
	if status == "ideal" {
		return 0
	}
	return 1
}

// isAllImunisasiIdeal checks if all immunizations for a baby are ideal.
// Returns true if all immunizations are ideal.
func isAllImunisasiIdeal(imunisasi Imunisasi) bool {
	statuses := []string{
		imunisasi.HB0.StatusImunisasi,
		imunisasi.BCG1.StatusImunisasi,
		imunisasi.Polio1.StatusImunisasi,
		imunisasi.Polio2.StatusImunisasi,
		imunisasi.Polio3.StatusImunisasi,
		imunisasi.Polio4.StatusImunisasi,
		imunisasi.DPTHbHib1.StatusImunisasi,
		imunisasi.DPTHbHib2.StatusImunisasi,
		imunisasi.DPTHbHib3.StatusImunisasi,
		imunisasi.IPV1.StatusImunisasi,
		imunisasi.IPV2.StatusImunisasi,
		imunisasi.ROTA1.StatusImunisasi,
		imunisasi.ROTA2.StatusImunisasi,
		imunisasi.ROTA3.StatusImunisasi,
		imunisasi.PCV1.StatusImunisasi,
		imunisasi.PCV2.StatusImunisasi,
		imunisasi.JE1.StatusImunisasi,
		imunisasi.MR1.StatusImunisasi,
		imunisasi.IDL1.StatusImunisasi,
	}

	for _, status := range statuses {
		if status != "ideal" {
			return false
		}
	}
	return true
}

// isPemberiImunivasiInvalid checks if the provider for immunization is invalid.
// Returns true if pemberi imunisasi is "invalid".
func isPemberiImunivasiInvalid(imunisasi Imunisasi) bool {
	pemberiImunisasiList := []string{
		imunisasi.HB0.PkmPemberiImunisasi,
		imunisasi.BCG1.PkmPemberiImunisasi,
		imunisasi.Polio1.PkmPemberiImunisasi,
		imunisasi.Polio2.PkmPemberiImunisasi,
		imunisasi.Polio3.PkmPemberiImunisasi,
		imunisasi.Polio4.PkmPemberiImunisasi,
		imunisasi.DPTHbHib1.PkmPemberiImunisasi,
		imunisasi.DPTHbHib2.PkmPemberiImunisasi,
		imunisasi.DPTHbHib3.PkmPemberiImunisasi,
		imunisasi.IPV1.PkmPemberiImunisasi,
		imunisasi.IPV2.PkmPemberiImunisasi,
		imunisasi.ROTA1.PkmPemberiImunisasi,
		imunisasi.ROTA2.PkmPemberiImunisasi,
		imunisasi.ROTA3.PkmPemberiImunisasi,
		imunisasi.PCV1.PkmPemberiImunisasi,
		imunisasi.PCV2.PkmPemberiImunisasi,
		imunisasi.JE1.PkmPemberiImunisasi,
		imunisasi.MR1.PkmPemberiImunisasi,
		imunisasi.IDL1.PkmPemberiImunisasi,
	}

	for _, pemberiImunisasi := range pemberiImunisasiList {
		pemberi := strings.ToLower(pemberiImunisasi)
		if pemberi == "-" {
			continue
		}
		if strings.Contains(pemberi, "wanasari") {
			return false
		}
	}
	return true
}

// isPosImunivasiInvalid checks if pos imunisasi is invalid.
// Returns true if pos imunisasi name is "invalid".
func isPosImunivasiInvalid(imunisasi Imunisasi) bool {
	posImunisasiList := []string{
		imunisasi.HB0.PosImunisasi,
		imunisasi.BCG1.PosImunisasi,
		imunisasi.Polio1.PosImunisasi,
		imunisasi.Polio2.PosImunisasi,
		imunisasi.Polio3.PosImunisasi,
		imunisasi.Polio4.PosImunisasi,
		imunisasi.DPTHbHib1.PosImunisasi,
		imunisasi.DPTHbHib2.PosImunisasi,
		imunisasi.DPTHbHib3.PosImunisasi,
		imunisasi.IPV1.PosImunisasi,
		imunisasi.IPV2.PosImunisasi,
		imunisasi.ROTA1.PosImunisasi,
		imunisasi.ROTA2.PosImunisasi,
		imunisasi.ROTA3.PosImunisasi,
		imunisasi.PCV1.PosImunisasi,
		imunisasi.PCV2.PosImunisasi,
		imunisasi.JE1.PosImunisasi,
		imunisasi.MR1.PosImunisasi,
		imunisasi.IDL1.PosImunisasi,
	}

	for _, posImunisasi := range posImunisasiList {
		pos := strings.ToLower(posImunisasi)
		if pos == "-" {
			continue
		}
		if strings.Contains(pos, "wanasari") {
			return false
		}
		if strings.Contains(pos, "dalam gedung") {
			return false
		}
		if strings.Contains(pos, "oleh sistem") {
			return false
		}
	}
	return true
}

// sortByTanggalLahir sorts dataImunisasiBayiList by tanggal lahir asc
func sortByTanggalLahir(dataImunisasiBayiList []DataImunisasiBayi) {
	sort.Slice(dataImunisasiBayiList, func(i, j int) bool {
		dateFormat := "2006-01-02"
		dateI, errI := time.Parse(dateFormat, dataImunisasiBayiList[i].TanggalLahirAnak)
		if errI != nil {
			log.Printf("Error parsing date for index %d: %v", i, errI)
			return false
		}
		dateJ, errJ := time.Parse(dateFormat, dataImunisasiBayiList[j].TanggalLahirAnak)
		if errJ != nil {
			log.Printf("Error parsing date for index %d: %v", j, errJ)
			return false
		}
		return dateI.Before(dateJ)
	})
}

const blackColor = "#000000"

// setStyle sets style for header
func setStyle(file *excelize.File, isHeader bool) (int, error) {

	bold := false
	if isHeader {
		bold = true
	}

	headerStyle, err := file.NewStyle(
		&excelize.Style{
			Font: &excelize.Font{
				Size:      12,
				Bold:      bold,
				Color:     blackColor,
				VertAlign: "center",
			},
			Alignment: &excelize.Alignment{
				Horizontal: "center",
			},
			Border: []excelize.Border{
				{
					Type:  "left",
					Style: 1,
					Color: blackColor,
				},
				{
					Type:  "right",
					Style: 1,
					Color: blackColor,
				},
				{
					Type:  "top",
					Style: 1,
					Color: blackColor,
				},
				{
					Type:  "bottom",
					Style: 1,
					Color: blackColor,
				},
			},
		},
	)

	if err != nil {
		log.Printf("Error creating style for header: %v", err)
		return 0, err
	}

	return headerStyle, err
}

// createNewExcelFile creates new excel file with given sheetName as default sheet name
func createNewExcelFile(sheetName string) (*excelize.File, error) {
	file := excelize.NewFile()
	index, err := file.NewSheet(sheetName)
	if err != nil {
		return nil, err
	}

	file.SetActiveSheet(index)
	file.SetDefaultFont("Times New Roman")

	return file, nil
}

// createNewSasaranImunisasiKejarBayiFile generates an Excel file containing the follow-up
// immunization data for babies who have missed some vaccinations.
func createNewSasaranImunisasiKejarBayiFile(dataImunisasiBayiList []DataImunisasiBayi) (*excelize.File, error) {
	sheetName := "Sheet1"
	file, err := createNewExcelFile(sheetName)
	if err != nil {
		log.Printf("Error creating sheet for new filtered data asik bayi file: %v", err)
		return nil, err
	}

	headers := []string{
		headerNamaAnak, headerUsiaAnak, headerTanggalLahirAnak, headerJenisKelaminAnak,
		headerNamaOrangTua, headerHBO, headerBCG1, headerPOLIO1, headerPOLIO2,
		headerPOLIO3, headerPOLIO4, headerDPTHbHib1, headerDPTHbHib2,
		headerDPTHbHib3, headerIPV1, headerIPV2, headerROTA1, headerROTA2,
		headerROTA3, headerPCV1, headerPCV2, headerJE1, headerMR1, headerIDL1,
	}
	headersMap := make(map[string]string)

	startHeaderRowNumAt := 3
	startHeaderRowAt := strconv.Itoa(startHeaderRowNumAt)

	headersLength := len(headers)
	columnNameList := columnName(headersLength)
	for i, columnName := range columnNameList {
		headerCell := columnName + startHeaderRowAt
		file.SetCellValue(sheetName, headerCell, headers[i])
		headersMap[headers[i]] = columnName
	}

	headerStyle, err := setStyle(file, true)
	if err == nil {
		firstHeaderCell := headersMap[headerNamaAnak] + startHeaderRowAt
		lastHeaderCell := headersMap[headerIDL1] + startHeaderRowAt
		file.SetCellStyle(sheetName, firstHeaderCell, lastHeaderCell, headerStyle)
	}

	title := strings.Replace(getSasaranImunisasiKejarBayiFileName(), ".xlsx", "", 1)
	titleStyle, err := file.NewStyle(
		&excelize.Style{
			Font: &excelize.Font{
				Size:      22,
				Bold:      true,
				Color:     blackColor,
				VertAlign: "center",
			},
			Alignment: &excelize.Alignment{
				Horizontal: "center",
			},
		},
	)
	firstRow := strconv.Itoa(1)
	file.SetCellValue(sheetName, headersMap[headerNamaAnak]+firstRow, title)
	file.MergeCell(sheetName, headersMap[headerNamaAnak]+firstRow, headersMap[headerIDL1]+firstRow)
	if err == nil {
		file.SetCellStyle(sheetName, headersMap[headerNamaAnak]+"1", headersMap[headerIDL1]+firstRow, titleStyle)
	}

	sortByTanggalLahir(dataImunisasiBayiList)

	startBodyRowNumAt := startHeaderRowNumAt + 1
	sasaranImunisasiKejarBayiList := newSasaranImunisasiKejarBayiList(dataImunisasiBayiList)
	for j, sasaranImunisasiKejarBayi := range sasaranImunisasiKejarBayiList {
		rowAt := strconv.Itoa(j + startBodyRowNumAt)
		file.SetCellValue(sheetName, headersMap[headerNamaAnak]+rowAt, sasaranImunisasiKejarBayi.NamaAnak)
		file.SetCellValue(sheetName, headersMap[headerUsiaAnak]+rowAt, sasaranImunisasiKejarBayi.UsiaAnak)
		file.SetCellValue(sheetName, headersMap[headerTanggalLahirAnak]+rowAt, sasaranImunisasiKejarBayi.TanggalLahirAnak)
		file.SetCellValue(sheetName, headersMap[headerJenisKelaminAnak]+rowAt, sasaranImunisasiKejarBayi.JenisKelaminAnak)
		file.SetCellValue(sheetName, headersMap[headerNamaOrangTua]+rowAt, sasaranImunisasiKejarBayi.NamaOrangTua)
		file.SetCellValue(sheetName, headersMap[headerHBO]+rowAt, sasaranImunisasiKejarBayi.HBO)
		file.SetCellValue(sheetName, headersMap[headerBCG1]+rowAt, sasaranImunisasiKejarBayi.BCG1)
		file.SetCellValue(sheetName, headersMap[headerPOLIO1]+rowAt, sasaranImunisasiKejarBayi.POLIO1)
		file.SetCellValue(sheetName, headersMap[headerPOLIO2]+rowAt, sasaranImunisasiKejarBayi.POLIO2)
		file.SetCellValue(sheetName, headersMap[headerPOLIO3]+rowAt, sasaranImunisasiKejarBayi.POLIO3)
		file.SetCellValue(sheetName, headersMap[headerPOLIO4]+rowAt, sasaranImunisasiKejarBayi.POLIO4)
		file.SetCellValue(sheetName, headersMap[headerDPTHbHib1]+rowAt, sasaranImunisasiKejarBayi.DPTHbHib1)
		file.SetCellValue(sheetName, headersMap[headerDPTHbHib2]+rowAt, sasaranImunisasiKejarBayi.DPTHbHib2)
		file.SetCellValue(sheetName, headersMap[headerDPTHbHib3]+rowAt, sasaranImunisasiKejarBayi.DPTHbHib3)
		file.SetCellValue(sheetName, headersMap[headerIPV1]+rowAt, sasaranImunisasiKejarBayi.IPV1)
		file.SetCellValue(sheetName, headersMap[headerIPV2]+rowAt, sasaranImunisasiKejarBayi.IPV2)
		file.SetCellValue(sheetName, headersMap[headerROTA1]+rowAt, sasaranImunisasiKejarBayi.ROTA1)
		file.SetCellValue(sheetName, headersMap[headerROTA2]+rowAt, sasaranImunisasiKejarBayi.ROTA2)
		file.SetCellValue(sheetName, headersMap[headerROTA3]+rowAt, sasaranImunisasiKejarBayi.ROTA3)
		file.SetCellValue(sheetName, headersMap[headerPCV1]+rowAt, sasaranImunisasiKejarBayi.PCV1)
		file.SetCellValue(sheetName, headersMap[headerPCV2]+rowAt, sasaranImunisasiKejarBayi.PCV2)
		file.SetCellValue(sheetName, headersMap[headerJE1]+rowAt, sasaranImunisasiKejarBayi.JE1)
		file.SetCellValue(sheetName, headersMap[headerMR1]+rowAt, sasaranImunisasiKejarBayi.MR1)
		file.SetCellValue(sheetName, headersMap[headerIDL1]+rowAt, sasaranImunisasiKejarBayi.IDL1)
	}

	bodyStyle, err := setStyle(file, false)
	if err == nil {
		firstBodyCell := headersMap[headerNamaAnak] + strconv.Itoa(startBodyRowNumAt)
		lastBodyCell := headersMap[headerIDL1] + strconv.Itoa(len(sasaranImunisasiKejarBayiList)+startHeaderRowNumAt)
		file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, bodyStyle)
	}

	for header, width := range getSasaranImunisasiKejarBayiColumnWidth() {
		file.SetColWidth(sheetName, headersMap[header], headersMap[header], width)
	}

	return file, nil
}

// getSasaranImunisasiKejarBayiColumnWidth returns column width for data kejar asik bayi
func getSasaranImunisasiKejarBayiColumnWidth() map[string]float64 {
	return map[string]float64{
		headerNamaAnak:         47,
		headerUsiaAnak:         15,
		headerTanggalLahirAnak: 20,
		headerJenisKelaminAnak: 20,
		headerNamaOrangTua:     40,
		headerHBO:              6,
		headerBCG1:             8,
		headerPOLIO1:           10,
		headerPOLIO2:           10,
		headerPOLIO3:           10,
		headerPOLIO4:           10,
		headerDPTHbHib1:        16,
		headerDPTHbHib2:        16,
		headerDPTHbHib3:        16,
		headerIPV1:             8,
		headerIPV2:             8,
		headerROTA1:            10,
		headerROTA2:            10,
		headerROTA3:            10,
		headerPCV1:             8,
		headerPCV2:             8,
		headerJE1:              8,
		headerMR1:              6,
		headerIDL1:             8,
	}
}

// sasaran imunisasi kejar bayi header const
const (
	headerNamaAnak         = "Nama Anak"
	headerUsiaAnak         = "Usia Anak"
	headerTanggalLahirAnak = "Tanggal Lahir Anak"
	headerJenisKelaminAnak = "Jenis Kelamin Anak"
	headerNamaOrangTua     = "Nama Orang Tua"
	headerHBO              = "HBO"
	headerBCG1             = "BCG 1"
	headerPOLIO1           = "POLIO 1"
	headerPOLIO2           = "POLIO 2"
	headerPOLIO3           = "POLIO 3"
	headerPOLIO4           = "POLIO 4"
	headerDPTHbHib1        = "DPT-Hb-Hib 1"
	headerDPTHbHib2        = "DPT-Hb-Hib 2"
	headerDPTHbHib3        = "DPT-Hb-Hib 3"
	headerIPV1             = "IPV 1"
	headerIPV2             = "IPV 2"
	headerROTA1            = "ROTA 1"
	headerROTA2            = "ROTA 2"
	headerROTA3            = "ROTA 3"
	headerPCV1             = "PCV 1"
	headerPCV2             = "PCV 2"
	headerJE1              = "JE 1"
	headerMR1              = "MR1"
	headerIDL1             = "IDL1"
)
