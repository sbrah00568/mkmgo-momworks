package imunisasi

import (
	"context"
	"fmt"
	"log"
	"sort"
	"strconv"
	"strings"
	"time"
)

// Define a key type for context
type contextKey string

// Create a key for the sasaranType value
const sasaranTypeKey contextKey = "sasaranType"

// GetSasaranTypeFromContext retrieves sasaran type from context
func GetSasaranTypeFromContext(ctx context.Context) string {
	if sasaranType, ok := ctx.Value(sasaranTypeKey).(string); ok {
		return sasaranType
	}
	return EMPTY_STRING
}

// common consts
const (
	EMPTY_STRING           = ""
	SPACE                  = " "
	IDL_1                  = "IDL 1"
	IBL_1                  = "IBL 1"
	BAYI                   = "bayi"
	NAMA_ANAK              = "Nama Anak"
	USIA_ANAK              = "Usia Anak"
	TANGGAL_LAHIR_ANAK     = "Tanggal Lahir Anak"
	JENIS_KELAMIN_ANAK     = "Jenis Kelamin Anak"
	NAMA_ORANG_TUA         = "Nama Orang Tua"
	PUSKESMAS              = "Puskesmas"
	TANGGAL                = "Tanggal"
	POS                    = "Pos"
	STATUS                 = "Status"
	A                      = "A"
	HYPHEN                 = "-"
	STATUS_IDL_1           = "Status IDL 1"
	STATUS_IMUNISASI_PCV_3 = "Status Imunisasi PCV 3"
)

// SasaranImunisasiConfig holds apps configuration for sasaran imunisasi
type SasaranImunisasiConfig struct {
	ColumnName             []string `yaml:"column_name"`
	DetailImunisasi        []string `yaml:"detail_imunisasi"`
	DetailImunisasiLengkap []string `yaml:"detail_imunisasi_lengkap"`
	ImunisasiBayi          []string `yaml:"imunisasi_bayi"`
	ImunisasiBaduta        []string `yaml:"imunisasi_baduta"`
}

// Sasaran represents sasaran imunisasi for both bayi and baduta
type SasaranImunisasi struct {
	NamaAnak         string                     `json:"namaAnak"`
	UsiaAnak         string                     `json:"usiaAnak"`
	TanggalLahirAnak string                     `json:"tanggalLahirAnak"`
	JenisKelaminAnak string                     `json:"jenisKelaminAnak"`
	NamaOrangTua     string                     `json:"namaOrangTua"`
	Puskesmas        string                     `json:"puskesmas"`
	DetailImunisasi  map[string]DetailImunisasi `json:"detailImunisasi"`
}

// DetailImunisasi represents detailed immunization data. Status ideal = 0 or non-ideal = 1
// Non-ideal means the recipient has not yet received the immunization
type DetailImunisasi struct {
	Tanggal map[string]string
	Pos     map[string]string
	Status  map[string]int
}

// SasaranImunisasiService manages sasaran imunisasi data and column mapping for bayi and baduta
type SasaranImunisasiService struct {
	Cfg                    *SasaranImunisasiConfig
	SourceFileColumnMap    map[string]Column
	SasaranBayiColumnMap   map[string]Column // represents xlsx column map for the generated file
	SasaranBadutaColumnMap map[string]Column // represents xlsx column map for the generated file
	SasaranImunisasiList   []SasaranImunisasi
}

// NewSasaranImunisasiService initializes a new instance of SasaranImunisasiService
// with column mappings for bayi and baduta based on the given config.
func NewSasaranImunisasiService(cfg *SasaranImunisasiConfig) *SasaranImunisasiService {
	sasaranBayiColumnMap := SetColumnMap(cfg, cfg.ImunisasiBayi)
	sasaranBadutaColumnMap := SetColumnMap(cfg, cfg.ImunisasiBaduta)
	return &SasaranImunisasiService{
		Cfg:                    cfg,
		SasaranBayiColumnMap:   sasaranBayiColumnMap,
		SasaranBadutaColumnMap: sasaranBadutaColumnMap,
	}
}

// SetColumnMap generates a map of column names to Column structures for the
// given set of immunization data (imunisasi).
func SetColumnMap(cfg *SasaranImunisasiConfig, imunisasi []string) map[string]Column {
	baseColumns := cfg.ColumnName
	detailImunisasi := cfg.DetailImunisasi
	detailImunisasiLengkap := cfg.DetailImunisasiLengkap

	columnMap := make(map[string]Column)
	colIndex := 1

	// Add base columns
	for _, columnName := range baseColumns {
		columnMap[columnName] = Column{Label: GetXlsxColumnLabel(colIndex)}
		colIndex++
	}

	// Add immunization columns
	for _, imun := range imunisasi {
		details := detailImunisasi
		if imun == IDL_1 || imun == IBL_1 {
			details = detailImunisasiLengkap
		}
		for _, detail := range details {
			columnMap[detail+SPACE+imun] = Column{Label: GetXlsxColumnLabel(colIndex)}
			colIndex++
		}
	}

	return columnMap
}

// GenerateFile processes the provided source Excel file and generates a new xlsx file
// based on the sasaran imunisasi data and column mappings.
// Returns a pointer to the generated xlsx file and an error if the generation fails.
func (svc *SasaranImunisasiService) GenerateFile(sourceFile XlsxSourceFile) (*XlsxGeneratedFile, error) {
	sasaranImunisasiList := []SasaranImunisasi{} // initialize sasaran imunisasi list

	// retrieves column map
	sasaranColumnMap, _ := svc.GetSasaranColumnMap(sourceFile.Ctx)
	sourceColumnMap := svc.GetSourceColumnMap(sourceFile)

	rowIndex := 2
	for {
		// check end of file
		if GetCellValue(sourceFile, A+strconv.Itoa(rowIndex)) == HYPHEN {
			break
		}

		// populate each rows data
		isRowValid, sasaranImunisasi := svc.PopulateRowsData(&DataRowPopulator{
			SasaranColumnMap: sasaranColumnMap,
			SourceColumnMap:  sourceColumnMap,
			RowIndex:         rowIndex,
			SourceFile:       sourceFile,
		})
		rowIndex++

		if isRowValid && sasaranImunisasi.CountNonIdealImmunizations() > 0 {
			sasaranImunisasiList = append(sasaranImunisasiList, sasaranImunisasi)
		}
	}

	// sort sasaran imunisasi anak by tanggal lahir from the oldest to the youngest
	SortSasaranImunisasiAnak(sasaranImunisasiList, func(s SasaranImunisasi) string {
		return s.TanggalLahirAnak
	})
	svc.SasaranImunisasiList = sasaranImunisasiList

	// create new xlsx file containing filtered data from source
	excelFile, err := CreateNewXlsxFile(sourceFile.Ctx, svc)
	if err != nil {
		return nil, err
	}

	return &XlsxGeneratedFile{
		FileName:     svc.GetFileName(GetSasaranTypeFromContext(sourceFile.Ctx)) + ".xlsx",
		ExcelizeFile: excelFile,
	}, nil
}

// DataRowPopulator contains the necessary information to populate a row of data into a SasaranImunisasi struct from a source Excel file.
type DataRowPopulator struct {
	SasaranColumnMap map[string]Column
	SourceColumnMap  map[string]Column
	RowIndex         int
	SourceFile       XlsxSourceFile
}

// GetSasaranColumnMap returns sasaran column map and last column label based on sasaran type
func (svc *SasaranImunisasiService) GetSasaranColumnMap(ctx context.Context) (map[string]Column, string) {
	sasaranType := GetSasaranTypeFromContext(ctx)
	if sasaranType == BAYI {
		return svc.SasaranBayiColumnMap, svc.SasaranBayiColumnMap[STATUS_IDL_1].Label
	}
	return svc.SasaranBadutaColumnMap, svc.SasaranBadutaColumnMap[STATUS_IMUNISASI_PCV_3].Label
}

// GetSourceColumnMap returns source column map which includes name and label (e.g.: name "ID" and label "A")
func (svc *SasaranImunisasiService) GetSourceColumnMap(sourceFile XlsxSourceFile) map[string]Column {
	sourceColumnMap := make(map[string]Column)

	// retrieves source label
	sourceCellIndex := 1
	for {
		sourceLabel := GetXlsxColumnLabel(sourceCellIndex)
		sourceCell := GetCellValue(sourceFile, sourceLabel+strconv.Itoa(1))
		if sourceCell == HYPHEN {
			break
		}
		sourceColumnMap[sourceCell] = Column{Label: sourceLabel}
		sourceCellIndex++
	}

	return sourceColumnMap
}

// PopulateRowsData populates the SasaranImunisasi struct with data from the specified row in the source file.
// It takes a DataRowPopulator which contains information about the row being processed, including
// mappings of column names and the source file itself. The method returns a boolean indicating
// whether the row data is valid and a pointer to the populated SasaranImunisasi struct.
func (svc *SasaranImunisasiService) PopulateRowsData(populator *DataRowPopulator) (bool, SasaranImunisasi) {
	isRowValid := true
	sasaranImunisasi := SasaranImunisasi{}
	strRowIndex := strconv.Itoa(populator.RowIndex)
	for sasaranColumnName := range populator.SasaranColumnMap {
		if sasaranColumnName == USIA_ANAK {
			continue
		}

		cell := populator.SourceColumnMap[sasaranColumnName].Label + strRowIndex
		sourceFile := populator.SourceFile

		if IsSourceInvalid(sasaranColumnName, cell, sourceFile) {
			isRowValid = false
			break
		}

		sasaranImunisasi.PopulateSasaranImunisasi(GetCellValue(sourceFile, cell), sasaranColumnName, svc.Cfg)
	}

	return isRowValid, sasaranImunisasi
}

// PopulateSasaranImunisasi populates sasaran imunisasi data for each column name with given cell value
func (sasaranImunisasi *SasaranImunisasi) PopulateSasaranImunisasi(cellValue, sasaranColumnName string, cfg *SasaranImunisasiConfig) {
	switch sasaranColumnName {
	case NAMA_ANAK:
		sasaranImunisasi.NamaAnak = cellValue
	case TANGGAL_LAHIR_ANAK:
		sasaranImunisasi.TanggalLahirAnak = cellValue
	case JENIS_KELAMIN_ANAK:
		sasaranImunisasi.JenisKelaminAnak = cellValue
	case NAMA_ORANG_TUA:
		sasaranImunisasi.NamaOrangTua = cellValue
	case PUSKESMAS:
		sasaranImunisasi.Puskesmas = cellValue
	default:
		detailImunisasi := sasaranImunisasi.GetDetailImunisasi(sasaranColumnName, cfg)
		switch {
		case strings.Contains(sasaranColumnName, TANGGAL):
			detailImunisasi.Tanggal[sasaranColumnName] = cellValue
		case strings.Contains(sasaranColumnName, POS):
			detailImunisasi.Pos[sasaranColumnName] = cellValue
		case strings.Contains(sasaranColumnName, STATUS):
			detailImunisasi.Status[sasaranColumnName] = GetStatusImunisasi(cellValue)
		}
	}

	if sasaranColumnName == TANGGAL_LAHIR_ANAK {
		sasaranImunisasi.UsiaAnak = sasaranImunisasi.CalculateUsiaAnak()
	}
}

// GetDetailImunisasi returns detail imunisasi of given sasaran imunisasi based on imunisasi type
func (s *SasaranImunisasi) GetDetailImunisasi(sasaranColumnName string, cfg *SasaranImunisasiConfig) DetailImunisasi {
	imunisasiType := s.GetImunisasiType(sasaranColumnName, cfg)
	if s.DetailImunisasi == nil {
		s.DetailImunisasi = make(map[string]DetailImunisasi)
	}

	if _, exists := s.DetailImunisasi[imunisasiType]; !exists {
		s.DetailImunisasi[imunisasiType] = DetailImunisasi{
			Tanggal: make(map[string]string),
			Pos:     make(map[string]string),
			Status:  make(map[string]int),
		}
	}

	return s.DetailImunisasi[imunisasiType]
}

// GetImunisasiType returns imunisasi type, for example if the column are Sasaran Imunisasi PCV 2 or Pos Imunisasi PCV 2 then will return "PCV 2"
func (sasaranImunisasi *SasaranImunisasi) GetImunisasiType(sasaranColumnName string, cfg *SasaranImunisasiConfig) string {
	for _, imunisasiBayi := range cfg.ImunisasiBayi {
		if strings.Contains(sasaranColumnName, imunisasiBayi) {
			return imunisasiBayi
		}
	}

	for _, imunisasiBaduta := range cfg.ImunisasiBaduta {
		if strings.Contains(sasaranColumnName, imunisasiBaduta) {
			return imunisasiBaduta
		}
	}

	return EMPTY_STRING
}

// GetStatusImunisasi returns an integer status based on the input string.
// "ideal" corresponds to 0, while any other status corresponds to 1.
func GetStatusImunisasi(status string) int {
	if status == "ideal" {
		return 0
	}
	return 1
}

// CountNonIdealImmunizations counts all non ideal imunisasi
func (s *SasaranImunisasi) CountNonIdealImmunizations() int {
	count := 0
	for _, detailImunisasi := range s.DetailImunisasi {
		for _, status := range detailImunisasi.Status {
			if status > 0 {
				count++
			}
		}
	}
	return count
}

// SortSasaranImunisasiAnak sorts data imunisasi anak based on tanggal lahir (DESC)
func SortSasaranImunisasiAnak[T any](list []T, dateExtractor func(T) string) {
	sort.Slice(list, func(i, j int) bool {
		dateFormat := "2006-01-02"
		parseDate := func(index int) (time.Time, error) {
			return time.Parse(dateFormat, dateExtractor(list[index]))
		}

		dateI, errI := parseDate(i)
		if errI != nil {
			log.Printf("Error parsing date for index %d: %v", i, errI)
			return false
		}

		dateJ, errJ := parseDate(j)
		if errJ != nil {
			log.Printf("Error parsing date for index %d: %v", j, errJ)
			return false
		}

		return dateI.Before(dateJ)
	})
}

// IsSourceInvalid checks whether the immunization information from source file is valid
// based on the given column and row index. It returns true if invalid, otherwise false.
func IsSourceInvalid(sasaranColumnName, cell string, sourceFile XlsxSourceFile) bool {
	if strings.Contains(sasaranColumnName, POS) {
		cellValue := strings.ToLower(GetCellValue(sourceFile, cell))
		switch {
		case strings.Contains(cellValue, "cibuntu"):
			return true
		case strings.Contains(cellValue, "wanasari"),
			strings.Contains(cellValue, "dalam gedung"),
			strings.Contains(cellValue, "oleh sistem"),
			cellValue == HYPHEN:
			return false
		}
	}
	return false
}

// CalculateUsiaAnak calculates the age of a child based on their birth date in the format "YYYY-MM-DD".
// It returns a string indicating the age in months and days.
func (sasaranImunisasi *SasaranImunisasi) CalculateUsiaAnak() string {
	birthDate, err := time.Parse("2006-01-02", sasaranImunisasi.TanggalLahirAnak)
	if err != nil {
		log.Printf("Failed to parse tanggal lahir anak: %v", err)
		return "-"
	}

	currentDate := time.Now()
	months := currentDate.Year()*12 + int(currentDate.Month()) - (birthDate.Year()*12 + int(birthDate.Month()))
	days := currentDate.Day() - birthDate.Day()

	if days < 0 {
		months--
		previousMonth := currentDate.AddDate(0, 0, -currentDate.Day()).Day()
		days += previousMonth // Add days from the previous month
	}

	return fmt.Sprintf("%d Bulan %d Hari", months, days)
}

// GetFileName returns title based on sasaranType
func (svc *SasaranImunisasiService) GetFileName(sasaranType string) string {
	return "Sasaran Imunisasi " + CapitalizeFirstChar(sasaranType) + SPACE + GetCurrentDateStr()
}

// CapitalizeFirstChar capitalizes the first character of a string.
func CapitalizeFirstChar(input string) string {
	if len(input) == 0 {
		return input
	}
	return strings.ToUpper(string(input[0])) + input[1:]
}

// GetCurrentDateStr returns the current Indonesian date as a formatted string in the format "Day Month" (e.g., "25 September").
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
	return fmt.Sprintf("%d %s", currentDate.Day(), months[currentDate.Month()])
}

// SetTitle sets the title of the Excel sheet for the generated file
func (svc *SasaranImunisasiService) SetTitle(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.TitleRowAt)
	sasaranImunisasiMap, lastColumnLabel := svc.GetSasaranColumnMap(newFile.Ctx)

	firstCell := sasaranImunisasiMap[NAMA_ANAK].Label + rowAt
	lastCell := lastColumnLabel + rowAt

	file.SetCellValue(sheetName, firstCell, svc.GetFileName(GetSasaranTypeFromContext(newFile.Ctx)))
	file.MergeCell(sheetName, firstCell, lastCell)
	file.SetCellStyle(sheetName, firstCell, lastCell, newFile.TitleStyle)
}

// SetHeader sets the header row of the Excel sheet
func (svc *SasaranImunisasiService) SetHeader(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.HeaderRowAt)
	sasaranImunisasiMap, lastColumnLabel := svc.GetSasaranColumnMap(newFile.Ctx)

	for name, column := range sasaranImunisasiMap {
		file.SetCellValue(sheetName, column.Label+rowAt, name)
	}

	firstBodyCell := sasaranImunisasiMap[NAMA_ANAK].Label + rowAt
	lastBodyCell := lastColumnLabel + rowAt

	file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.HeaderStyle)
}

// SetBody sets the body row of the Excel sheet
func (svc *SasaranImunisasiService) SetBody(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	sasaranImunisasiMap, lastColumnLabel := svc.GetSasaranColumnMap(newFile.Ctx)

	for i, sasaranImunisasi := range svc.SasaranImunisasiList {
		rowAt := strconv.Itoa(i + newFile.StartBodyRowAt)
		file.SetCellValue(sheetName, sasaranImunisasiMap[NAMA_ANAK].Label+rowAt, sasaranImunisasi.NamaAnak)
		file.SetCellValue(sheetName, sasaranImunisasiMap[USIA_ANAK].Label+rowAt, sasaranImunisasi.UsiaAnak)
		file.SetCellValue(sheetName, sasaranImunisasiMap[TANGGAL_LAHIR_ANAK].Label+rowAt, sasaranImunisasi.TanggalLahirAnak)
		file.SetCellValue(sheetName, sasaranImunisasiMap[JENIS_KELAMIN_ANAK].Label+rowAt, sasaranImunisasi.JenisKelaminAnak)
		file.SetCellValue(sheetName, sasaranImunisasiMap[NAMA_ORANG_TUA].Label+rowAt, sasaranImunisasi.NamaOrangTua)
		file.SetCellValue(sheetName, sasaranImunisasiMap[PUSKESMAS].Label+rowAt, sasaranImunisasi.Puskesmas)
		for _, detailImunisasi := range sasaranImunisasi.DetailImunisasi {
			for key, value := range detailImunisasi.Tanggal {
				cell := sasaranImunisasiMap[key].Label + rowAt
				file.SetCellValue(sheetName, cell, value)
			}

			for key, value := range detailImunisasi.Pos {
				cell := sasaranImunisasiMap[key].Label + rowAt
				file.SetCellValue(sheetName, cell, value)
			}

			for key, value := range detailImunisasi.Status {
				cell := sasaranImunisasiMap[key].Label + rowAt
				file.SetCellValue(sheetName, cell, value)
			}
		}

		firstBodyCell := sasaranImunisasiMap[NAMA_ANAK].Label + rowAt
		lastBodyCell := lastColumnLabel + rowAt
		file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.BodyStyle)
	}
}

// SetColumnWidth sets the column width of the Excel sheet
func (svc *SasaranImunisasiService) SetColumnWidth(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	sasaranImunisasiMap, _ := svc.GetSasaranColumnMap(newFile.Ctx)
	for _, column := range sasaranImunisasiMap {
		file.SetColWidth(sheetName, column.Label, column.Label, 32)
	}
}
