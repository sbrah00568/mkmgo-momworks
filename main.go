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
	"gopkg.in/yaml.v2"
)

type XlsxColumn struct {
	Code   string `yaml:"code"`
	Index  int    `yaml:"index"`
	Name   string `yaml:"name"`
	Label  string `yaml:"label"`
	Length int    `yaml:"length"`
}

type Config struct {
	ColumnDataImunisasiBayi    []XlsxColumn `yaml:"column_data_imunisasi_bayi"`
	ColumnDataImunaisasiBaduta []XlsxColumn `yaml:"column_data_imunisasi_baduta"`
	ColumnSasaranKejarBayi     []XlsxColumn `yaml:"column_sasaran_kejar_bayi"`
	ColumnSasaranKejarBaduta   []XlsxColumn `yaml:"column_sasaran_kejar_baduta"`
}

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
	cfg, err := LoadConfig()
	if err != nil {
		panic(err)
	}

	sasaranKejarHandler := NewSasaranKejarHandler(NewKejarBayiService(cfg), NewKejarBadutaService(cfg))

	http.HandleFunc("/momworks/sasaran/kejar", sasaranKejarHandler.GenerateFileHandler)
	log.Println("Starting momworks server....")
	if err := http.ListenAndServe("localhost:8080", nil); err != nil {
		log.Fatalf("Server failed: %v", err)
	}
}

type SourceXlsxFile struct {
	TempFilePath string
	SheetName    string
	ExcelizeFile *excelize.File
}

type XlsxGeneratedFile struct {
	FileName     string
	ExcelizeFile *excelize.File
}

type XlsxFileTransformer interface {
	GenerateFile(sourceFile SourceXlsxFile) (*XlsxGeneratedFile, error)
}

type SasaranKejarHandler struct {
	KejarBayiService   XlsxFileTransformer
	KejarBadutaService XlsxFileTransformer
}

func NewSasaranKejarHandler(bayiSvc, badutaSvc XlsxFileTransformer) *SasaranKejarHandler {
	return &SasaranKejarHandler{
		KejarBayiService:   bayiSvc,
		KejarBadutaService: badutaSvc,
	}
}

func (h *SasaranKejarHandler) GenerateFileHandler(w http.ResponseWriter, r *http.Request) {
	const maxUploadSize = 10 << 20

	if err := r.ParseMultipartForm(maxUploadSize); err != nil {
		http.Error(w, "File size too large", http.StatusBadRequest)
		log.Printf("File upload error: %v", err)
		return
	}

	src, fileHeader, err := r.FormFile("myFile")
	if err != nil {
		http.Error(w, "Failed to retrieve file", http.StatusBadRequest)
		log.Printf("Error retrieving file: %v", err)
		return
	}
	defer src.Close()

	log.Printf("Uploaded File: %s, Size: %d, MIME: %v", fileHeader.Filename, fileHeader.Size, fileHeader.Header)

	tempFilePath, err := SaveFileToTemp(src, fileHeader.Filename)
	if err != nil {
		http.Error(w, "Failed to save file", http.StatusInternalServerError)
		log.Printf("Error saving file: %v", err)
		return
	}
	defer os.Remove(tempFilePath)

	sheetName := r.FormValue("sheetName")

	excelizeFile, err := OpenFile(tempFilePath)
	if err != nil {
		http.Error(w, "Failed to open file", http.StatusInternalServerError)
		log.Printf("Error opening Excel file: %v", err)
		return
	}
	defer excelizeFile.Close()

	var service XlsxFileTransformer
	switch r.FormValue("type") {
	case "bayi":
		service = h.KejarBayiService
	case "baduta":
		service = h.KejarBadutaService
	default:
		http.Error(w, "invalid sasaran type. Must be either bayi or baduta", http.StatusBadRequest)
		log.Printf("Invalid sasaran type: %s", r.FormValue("type"))
		return
	}

	sourceFile := SourceXlsxFile{
		TempFilePath: tempFilePath,
		SheetName:    sheetName,
		ExcelizeFile: excelizeFile,
	}
	generatedFile, err := service.GenerateFile(sourceFile)
	if err != nil {
		http.Error(w, "Error creating file", http.StatusInternalServerError)
		log.Printf("Error generating file: %v", err)
		return
	}

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", `attachment; filename="`+generatedFile.FileName+`"`)

	if err := generatedFile.ExcelizeFile.Write(w); err != nil {
		http.Error(w, "Unable to generate file", http.StatusInternalServerError)
		return
	}

	log.Println("Successfully uploaded and processed file")
}

func SaveFileToTemp(file io.Reader, filename string) (string, error) {
	tempFile, err := os.CreateTemp("temp", filepath.Base(filename)+"-*.xlsx")
	if err != nil {
		log.Printf("Error creating temp file: %v", err)
		return "", fmt.Errorf("failed to create temp file: %w", err)
	}
	defer tempFile.Close()

	if _, err := io.Copy(tempFile, file); err != nil {
		log.Printf("Error writing to temp file: %v", err)
		return "", fmt.Errorf("failed to write to temp file: %w", err)
	}

	return tempFile.Name(), nil
}

func OpenFile(filePath string) (*excelize.File, error) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	return file, nil
}

type NewXlsxFile struct {
	SheetName      string
	ExcelizeFile   *excelize.File
	TitleRowAt     int
	HeaderRowAt    int
	StartBodyRowAt int
	TitleStyle     int
	HeaderStyle    int
	BodyStyle      int
}

type NewXlsxGenerator interface {
	SetTitle(newFile NewXlsxFile)
	SetHeader(newFile NewXlsxFile)
	SetBody(newFile NewXlsxFile)
	SetColumnWidth(newFile NewXlsxFile)
}

func CreateNewXlsxFile(generator NewXlsxGenerator) (*excelize.File, error) {
	excelizeFile := excelize.NewFile()
	sheetName := "Sheet1"
	index, err := excelizeFile.NewSheet(sheetName)
	if err != nil {
		return nil, err
	}

	excelizeFile.SetActiveSheet(index)
	excelizeFile.SetDefaultFont("Times New Roman")

	newXlsxFile := NewXlsxFile{
		SheetName:      sheetName,
		ExcelizeFile:   excelizeFile,
		TitleRowAt:     0,
		HeaderRowAt:    2,
		StartBodyRowAt: 3,
	}

	titleStyle, err := excelizeFile.NewStyle(
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
	if err != nil {
		titleStyle = 0
	}
	newXlsxFile.TitleStyle = titleStyle

	headerStyle := SetXlsxStyle(excelizeFile, true)
	newXlsxFile.HeaderStyle = headerStyle

	bodyStyle := SetXlsxStyle(excelizeFile, false)
	newXlsxFile.BodyStyle = bodyStyle

	generator.SetTitle(newXlsxFile)
	generator.SetHeader(newXlsxFile)
	generator.SetBody(newXlsxFile)
	generator.SetColumnWidth(newXlsxFile)

	return excelizeFile, nil
}

const blackColor = "#000000"

func SetXlsxStyle(file *excelize.File, isHeader bool) int {

	bold := false
	if isHeader {
		bold = true
	}

	style, err := file.NewStyle(
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
		return 0
	}

	return style
}

func GetCurrentDateStr() string {
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

	currentDate := time.Now()
	return strconv.Itoa(currentDate.Day()) + " " + months[currentDate.Month()]
}

type SasaranKejarBayi struct {
	NamaAnak         string `json:"namaAnak"`
	UsiaAnak         string `json:"usiaAnak"`
	TanggalLahirAnak string `json:"tanggalLahirAnak"`
	JenisKelaminAnak string `json:"jenisKelaminAnak"`
	NamaOrangTua     string `json:"namaOrangTua"`
	Puskesmas        string `json:"puskesmas"`
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

func (s *SasaranKejarBayi) SetCellValue(code, cell string, sourceFile SourceXlsxFile) {
	cellHandlers := map[string]func(){
		"nama_anak":                     func() { s.NamaAnak = GetCellValue(sourceFile, cell) },
		"tanggal_lahir_anak":            func() { s.TanggalLahirAnak = GetCellValue(sourceFile, cell) },
		"jenis_kelamin_anak":            func() { s.JenisKelaminAnak = GetCellValue(sourceFile, cell) },
		"nama_orang_tua":                func() { s.NamaOrangTua = GetCellValue(sourceFile, cell) },
		"status_imunisasi_hb0":          func() { s.HBO = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_bcg_1":        func() { s.BCG1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_polio_1":      func() { s.POLIO1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_polio_2":      func() { s.POLIO2 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_polio_3":      func() { s.POLIO3 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_polio_4":      func() { s.POLIO4 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_dpt_hb_hib_1": func() { s.DPTHbHib1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_dpt_hb_hib_2": func() { s.DPTHbHib2 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_dpt_hb_hib_3": func() { s.DPTHbHib3 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_ipv_1":        func() { s.IPV1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_ipv_2":        func() { s.IPV2 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_rota_1":       func() { s.ROTA1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_rota_2":       func() { s.ROTA2 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_rota_3":       func() { s.ROTA3 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_pcv_1":        func() { s.PCV1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_pcv_2":        func() { s.PCV2 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_je_1":         func() { s.JE1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_mr_1":         func() { s.MR1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_idl_1":                  func() { s.IDL1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
	}

	if handler, exists := cellHandlers[code]; exists {
		handler()
	}
}

func (s *SasaranKejarBayi) CountImunisasiTidakIdeal() int {
	imunisasiBayiList := []int{
		s.HBO,
		s.BCG1,
		s.POLIO1,
		s.POLIO2,
		s.POLIO3,
		s.POLIO4,
		s.DPTHbHib1,
		s.DPTHbHib2,
		s.DPTHbHib3,
		s.IPV1,
		s.IPV2,
		s.ROTA1,
		s.ROTA2,
		s.ROTA3,
		s.PCV1,
		s.PCV2,
		s.JE1,
		s.MR1,
		s.IDL1,
	}
	return SumListInt(imunisasiBayiList)
}

func (s *SasaranKejarBayi) PopulateSasaranKejar(svc *KejarBayiService, rowIndex int, sourceFile SourceXlsxFile) (bool, bool) {
	isNoMoreRow, isImunisasiBayiValid := false, false

	for _, column := range svc.DataImunisasiBayiMap {
		isNoMoreRow = CheckForEndOfFile(column, rowIndex, sourceFile)
		if isNoMoreRow {
			break
		}

		isImunisasiBayiValid = IsImunisasiValid(column, rowIndex, sourceFile)
		if isImunisasiBayiValid {
			s.SetCellValue(column.Code, column.Label+strconv.Itoa(rowIndex), sourceFile)
		}
	}

	return isNoMoreRow, isImunisasiBayiValid
}

func (s *SasaranKejarBayi) AppendIfValid(isImunisasiBayiValid bool, sasaranKejarBayiList []SasaranKejarBayi) []SasaranKejarBayi {
	if s.CountImunisasiTidakIdeal() != 0 && isImunisasiBayiValid {
		s.Puskesmas = "WANASARI"
		s.UsiaAnak = CalculateUsiaAnak(s.TanggalLahirAnak)
		return append(sasaranKejarBayiList, *s)
	}

	return sasaranKejarBayiList
}

type KejarBayiService struct {
	DataImunisasiBayiMap map[string]XlsxColumn
	SasaranKejarBayiMap  map[string]XlsxColumn
	SasaranKejarBayiList []SasaranKejarBayi
}

func NewKejarBayiService(config *Config) *KejarBayiService {
	return &KejarBayiService{
		DataImunisasiBayiMap: PopulateXlsxColumnMap(config.ColumnDataImunisasiBayi),
		SasaranKejarBayiMap:  PopulateXlsxColumnMap(config.ColumnSasaranKejarBayi),
	}
}

func (svc *KejarBayiService) GenerateFile(sourceFile SourceXlsxFile) (*XlsxGeneratedFile, error) {
	var sasaranKejarBayiList []SasaranKejarBayi
	rowIndex := 2
	for {
		sasaranKejarBayi := SasaranKejarBayi{}
		isNoMoreRow, isImunisasiBayiValid := sasaranKejarBayi.PopulateSasaranKejar(svc, rowIndex, sourceFile)

		if isNoMoreRow {
			break
		}

		sasaranKejarBayiList = sasaranKejarBayi.AppendIfValid(isImunisasiBayiValid, sasaranKejarBayiList)

		rowIndex++
	}

	SortSasaranKejarList(sasaranKejarBayiList, func(s SasaranKejarBayi) string {
		return s.TanggalLahirAnak
	})

	svc.SasaranKejarBayiList = sasaranKejarBayiList

	excelizeFile, err := CreateNewXlsxFile(svc)
	if err != nil {
		return nil, err
	}

	return &XlsxGeneratedFile{
		FileName:     "Sasaran Imunisasi Kejar Bayi " + GetCurrentDateStr() + ".xlsx",
		ExcelizeFile: excelizeFile,
	}, nil
}

func (svc *KejarBayiService) SetTitle(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.TitleRowAt)
	title := "Sasaran Imunisasi Kejar Bayi " + GetCurrentDateStr()

	firstCell := svc.SasaranKejarBayiMap["nama_anak"].Label + rowAt
	lastCell := svc.SasaranKejarBayiMap["status_idl_1"].Label + rowAt

	file.SetCellValue(sheetName, firstCell, title)
	file.MergeCell(sheetName, firstCell, lastCell)
	file.SetCellStyle(sheetName, firstCell, lastCell, newFile.TitleStyle)
}

func (svc *KejarBayiService) SetHeader(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.HeaderRowAt)

	for _, column := range svc.SasaranKejarBayiMap {
		file.SetCellValue(sheetName, column.Label+rowAt, column.Name)
	}

	firstBodyCell := svc.SasaranKejarBayiMap["nama_anak"].Label + rowAt
	lastBodyCell := svc.SasaranKejarBayiMap["status_idl_1"].Label + rowAt
	file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.HeaderStyle)
}

func (svc *KejarBayiService) SetBody(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	for i, sasaranKejarBayi := range svc.SasaranKejarBayiList {
		rowAt := strconv.Itoa(i + newFile.StartBodyRowAt)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["nama_anak"].Label+rowAt, sasaranKejarBayi.NamaAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["usia_anak"].Label+rowAt, sasaranKejarBayi.UsiaAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["tanggal_lahir_anak"].Label+rowAt, sasaranKejarBayi.TanggalLahirAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["jenis_kelamin_anak"].Label+rowAt, sasaranKejarBayi.JenisKelaminAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["nama_orang_tua"].Label+rowAt, sasaranKejarBayi.NamaOrangTua)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["puskesmas"].Label+rowAt, sasaranKejarBayi.Puskesmas)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_hb0"].Label+rowAt, sasaranKejarBayi.HBO)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_bcg_1"].Label+rowAt, sasaranKejarBayi.BCG1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_polio_1"].Label+rowAt, sasaranKejarBayi.POLIO1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_polio_2"].Label+rowAt, sasaranKejarBayi.POLIO2)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_polio_3"].Label+rowAt, sasaranKejarBayi.POLIO3)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_polio_4"].Label+rowAt, sasaranKejarBayi.POLIO4)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_dpt_hb_hib_1"].Label+rowAt, sasaranKejarBayi.DPTHbHib1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_dpt_hb_hib_2"].Label+rowAt, sasaranKejarBayi.DPTHbHib2)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_dpt_hb_hib_3"].Label+rowAt, sasaranKejarBayi.DPTHbHib3)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_ipv_1"].Label+rowAt, sasaranKejarBayi.IPV1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_ipv_2"].Label+rowAt, sasaranKejarBayi.IPV2)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_rota_1"].Label+rowAt, sasaranKejarBayi.ROTA1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_rota_2"].Label+rowAt, sasaranKejarBayi.ROTA2)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_rota_3"].Label+rowAt, sasaranKejarBayi.ROTA3)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_pcv_1"].Label+rowAt, sasaranKejarBayi.PCV1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_pcv_2"].Label+rowAt, sasaranKejarBayi.PCV2)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_je_1"].Label+rowAt, sasaranKejarBayi.JE1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_imunisasi_mr_1"].Label+rowAt, sasaranKejarBayi.MR1)
		file.SetCellValue(sheetName, svc.SasaranKejarBayiMap["status_idl_1"].Label+rowAt, sasaranKejarBayi.IDL1)
	}

	firstBodyCell := svc.SasaranKejarBayiMap["nama_anak"].Label + strconv.Itoa(newFile.StartBodyRowAt)
	lastBodyCell := svc.SasaranKejarBayiMap["status_idl_1"].Label + strconv.Itoa(len(svc.SasaranKejarBayiList)+newFile.HeaderRowAt)
	file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.BodyStyle)

}

func (svc *KejarBayiService) SetColumnWidth(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	for _, column := range svc.SasaranKejarBayiMap {
		file.SetColWidth(sheetName, column.Label, column.Label, float64(column.Length))
	}
}

type SasaranKejarBaduta struct {
	NamaAnak         string `json:"namaAnak"`
	UsiaAnak         string `json:"usiaAnak"`
	TanggalLahirAnak string `json:"tanggalLahirAnak"`
	JenisKelaminAnak string `json:"jenisKelaminAnak"`
	NamaOrangTua     string `json:"namaOrangTua"`
	Puskesmas        string `json:"puskesmas"`
	DPTHbHib4        int    `json:"dptHbHib4"`
	MR2              int    `json:"mr2"`
	IBL1             int    `json:"ibl1"`
	PCV3             int    `json:"pcv3"`
}

func (s *SasaranKejarBaduta) SetCellValue(code, cell string, sourceFile SourceXlsxFile) {
	cellHandlers := map[string]func(){
		"nama_anak":                     func() { s.NamaAnak = GetCellValue(sourceFile, cell) },
		"tanggal_lahir_anak":            func() { s.TanggalLahirAnak = GetCellValue(sourceFile, cell) },
		"jenis_kelamin_anak":            func() { s.JenisKelaminAnak = GetCellValue(sourceFile, cell) },
		"nama_orang_tua":                func() { s.NamaOrangTua = GetCellValue(sourceFile, cell) },
		"status_imunisasi_dpt_hb_hib_4": func() { s.DPTHbHib4 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_mr_2":         func() { s.MR2 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_ibl_1":                  func() { s.IBL1 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
		"status_imunisasi_pcv_3":        func() { s.PCV3 = GetStatusImunisasiKejar(GetCellValue(sourceFile, cell)) },
	}

	if handler, exists := cellHandlers[code]; exists {
		handler()
	}
}

func (s *SasaranKejarBaduta) CountImunisasiTidakIdeal() int {
	imunisasiBadutaList := []int{
		s.DPTHbHib4,
		s.MR2,
		s.IBL1,
		s.PCV3,
	}
	return SumListInt(imunisasiBadutaList)
}

func (s *SasaranKejarBaduta) PopulateSasaranKejar(svc *KejarBadutaService, rowIndex int, sourceFile SourceXlsxFile) (bool, bool) {
	isNoMoreRow, isImunisasiBayiValid := false, false

	for _, column := range svc.DataImunisasiBadutaMap {
		isNoMoreRow = CheckForEndOfFile(column, rowIndex, sourceFile)
		if isNoMoreRow {
			break
		}

		isImunisasiBayiValid = IsImunisasiValid(column, rowIndex, sourceFile)
		if isImunisasiBayiValid {
			s.SetCellValue(column.Code, column.Label+strconv.Itoa(rowIndex), sourceFile)
		}
	}

	return isNoMoreRow, isImunisasiBayiValid
}

func (s *SasaranKejarBaduta) AppendIfValid(isImunisasiBayiValid bool, sasaranKejarBadutaList []SasaranKejarBaduta) []SasaranKejarBaduta {
	if s.CountImunisasiTidakIdeal() != 0 && isImunisasiBayiValid {
		s.Puskesmas = "WANASARI"
		s.UsiaAnak = CalculateUsiaAnak(s.TanggalLahirAnak)
		return append(sasaranKejarBadutaList, *s)
	}

	return sasaranKejarBadutaList
}

type KejarBadutaService struct {
	DataImunisasiBadutaMap map[string]XlsxColumn
	SasaranKejarBadutaMap  map[string]XlsxColumn
	SasaranKejarBadutaList []SasaranKejarBaduta
}

func NewKejarBadutaService(config *Config) *KejarBadutaService {
	return &KejarBadutaService{
		DataImunisasiBadutaMap: PopulateXlsxColumnMap(config.ColumnDataImunaisasiBaduta),
		SasaranKejarBadutaMap:  PopulateXlsxColumnMap(config.ColumnSasaranKejarBaduta),
	}
}

func (svc *KejarBadutaService) GenerateFile(sourceFile SourceXlsxFile) (*XlsxGeneratedFile, error) {
	var sasaranKejarBadutaList []SasaranKejarBaduta

	rowIndex := 2
	for {
		sasaranKejarBaduta := SasaranKejarBaduta{}
		isNoMoreRow, isImunisasiBayiValid := sasaranKejarBaduta.PopulateSasaranKejar(svc, rowIndex, sourceFile)

		if isNoMoreRow {
			break
		}

		sasaranKejarBadutaList = sasaranKejarBaduta.AppendIfValid(isImunisasiBayiValid, sasaranKejarBadutaList)

		rowIndex++
	}

	SortSasaranKejarList(sasaranKejarBadutaList, func(s SasaranKejarBaduta) string {
		return s.TanggalLahirAnak
	})

	svc.SasaranKejarBadutaList = sasaranKejarBadutaList

	excelizeFile, err := CreateNewXlsxFile(svc)
	if err != nil {
		return nil, err
	}

	return &XlsxGeneratedFile{
		FileName:     "Sasaran Imunisasi Kejar Baduta " + GetCurrentDateStr() + ".xlsx",
		ExcelizeFile: excelizeFile,
	}, nil
}

func (svc *KejarBadutaService) SetTitle(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.TitleRowAt)
	title := "Sasaran Imunisasi Kejar Baduta " + GetCurrentDateStr()

	firstCell := svc.SasaranKejarBadutaMap["nama_anak"].Label + rowAt
	lastCell := svc.SasaranKejarBadutaMap["status_imunisasi_pcv_3"].Label + rowAt

	file.SetCellValue(sheetName, firstCell, title)
	file.MergeCell(sheetName, firstCell, lastCell)
	file.SetCellStyle(sheetName, firstCell, lastCell, newFile.TitleStyle)
}

func (svc *KejarBadutaService) SetHeader(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.HeaderRowAt)

	for _, column := range svc.SasaranKejarBadutaMap {
		file.SetCellValue(sheetName, column.Label+rowAt, column.Name)
	}

	firstBodyCell := svc.SasaranKejarBadutaMap["nama_anak"].Label + rowAt
	lastBodyCell := svc.SasaranKejarBadutaMap["status_imunisasi_pcv_3"].Label + rowAt
	file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.HeaderStyle)
}

func (svc *KejarBadutaService) SetBody(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	for i, sasaranKejarBaduta := range svc.SasaranKejarBadutaList {
		rowAt := strconv.Itoa(i + newFile.StartBodyRowAt)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["nama_anak"].Label+rowAt, sasaranKejarBaduta.NamaAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["usia_anak"].Label+rowAt, sasaranKejarBaduta.UsiaAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["tanggal_lahir_anak"].Label+rowAt, sasaranKejarBaduta.TanggalLahirAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["jenis_kelamin_anak"].Label+rowAt, sasaranKejarBaduta.JenisKelaminAnak)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["nama_orang_tua"].Label+rowAt, sasaranKejarBaduta.NamaOrangTua)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["puskesmas"].Label+rowAt, sasaranKejarBaduta.Puskesmas)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["status_imunisasi_dpt_hb_hib_4"].Label+rowAt, sasaranKejarBaduta.DPTHbHib4)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["status_imunisasi_mr_2"].Label+rowAt, sasaranKejarBaduta.MR2)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["status_ibl_1"].Label+rowAt, sasaranKejarBaduta.IBL1)
		file.SetCellValue(sheetName, svc.SasaranKejarBadutaMap["status_imunisasi_pcv_3"].Label+rowAt, sasaranKejarBaduta.PCV3)
	}

	firstBodyCell := svc.SasaranKejarBadutaMap["nama_anak"].Label + strconv.Itoa(newFile.StartBodyRowAt)
	lastBodyCell := svc.SasaranKejarBadutaMap["status_imunisasi_pcv_3"].Label + strconv.Itoa(len(svc.SasaranKejarBadutaList)+newFile.HeaderRowAt)
	file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.BodyStyle)
}

func (svc *KejarBadutaService) SetColumnWidth(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	for _, column := range svc.SasaranKejarBadutaMap {
		file.SetColWidth(sheetName, column.Label, column.Label, float64(column.Length))
	}
}

func CheckForEndOfFile(column XlsxColumn, rowIndex int, sourceFile SourceXlsxFile) bool {
	return column.Code == "id" && GetCellValue(sourceFile, column.Label+strconv.Itoa(rowIndex)) == "-"
}

func PopulateXlsxColumnMap(columns []XlsxColumn) map[string]XlsxColumn {
	columnMap := make(map[string]XlsxColumn, len(columns))
	for _, col := range columns {
		columnMap[col.Code] = col
	}
	return columnMap
}

func GetCellValue(sourceFile SourceXlsxFile, cell string) string {
	cellValue, err := sourceFile.ExcelizeFile.GetCellValue(sourceFile.SheetName, cell)
	if err != nil || cellValue == "" {
		cellValue = "-"
	}
	return cellValue
}

func SumListInt(values []int) int {
	total := 0
	for _, value := range values {
		total += value
	}
	return total
}

func CalculateUsiaAnak(tanggalLahirAnak string) string {
	birthDate, err := time.Parse("2006-01-02", tanggalLahirAnak)
	if err != nil {
		log.Printf("Failed to parse tanggal lahir anak: %v", err)
		return "-"
	}

	currentDate := time.Now()
	months := currentDate.Year()*12 + int(currentDate.Month()) - (birthDate.Year()*12 + int(birthDate.Month()))
	days := currentDate.Day() - birthDate.Day()

	if days < 0 {
		months--
		days += time.Date(currentDate.Year(), currentDate.Month(), 0, 0, 0, 0, 0, currentDate.Location()).Day()
	}

	return fmt.Sprintf("%d Bulan %d Hari", months, days)
}

func IsImunisasiValid(column XlsxColumn, rowIndex int, sourceFile SourceXlsxFile) bool {
	if !strings.Contains(column.Code, "_pemberi_imunisasi_") && !strings.Contains(column.Code, "pos_imunisasi_") {
		return true
	}

	cellValue := strings.ToLower(GetCellValue(sourceFile, column.Label+strconv.Itoa(rowIndex)))

	validValues := []string{"wanasari", "dalam gedung", "oleh sistem", "-"}
	for _, valid := range validValues {
		if strings.Contains(cellValue, valid) {
			return true
		}
	}

	return false
}

func GetStatusImunisasiKejar(status string) int {
	if status == "ideal" {
		return 0
	}
	return 1
}

func SortSasaranKejarList[T any](list []T, dateExtractor func(T) string) {
	sort.Slice(list, func(i, j int) bool {
		dateFormat := "2006-01-02"
		dateI, errI := time.Parse(dateFormat, dateExtractor(list[i]))
		if errI != nil {
			log.Printf("Error parsing date for index %d: %v", i, errI)
			return false
		}
		dateJ, errJ := time.Parse(dateFormat, dateExtractor(list[j]))
		if errJ != nil {
			log.Printf("Error parsing date for index %d: %v", j, errJ)
			return false
		}
		return dateI.Before(dateJ)
	})
}
