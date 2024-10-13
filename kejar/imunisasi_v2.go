package kejar

import (
	"context"
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"
)

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

// Column represents column characteristics for the generated sasaran imunisasi xlsx file
type Column struct {
	Label string
	Width float64
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

// GetXlsxColumnLabel generates an Excel column label based on a zero-based index.
// For example, index 0 returns "A", index 1 returns "B", and so on.
func GetXlsxColumnLabel(index int) string {
	if index == 0 {
		return "A"
	}

	label := EMPTY_STRING
	for index > 0 {
		index-- // Excel column index is 1-based, adjust to 0-based
		label = string(rune('A'+(index%26))) + label
		index /= 26
	}
	return label
}

// GenerateFile processes the provided source Excel file and generates a new xlsx file
// based on the sasaran imunisasi data and column mappings.
// Returns a pointer to the generated xlsx file and an error if the generation fails.
func (svc *SasaranImunisasiService) GenerateFile(sourceFile SourceXlsxFile) (*XlsxGeneratedFile, error) {
	sasaranImunisasiList := []SasaranImunisasi{}

	sasaranColumnMap, _ := svc.GetSasaranColumnMap(sourceFile.Ctx)

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

	rowIndex := 2
	for {
		isRowValid := true

		// check end of file
		if GetCellValue(sourceFile, A+strconv.Itoa(rowIndex)) == HYPHEN {
			break
		}

		// populate each rows data
		strRowIndex := strconv.Itoa(rowIndex)
		sasaranImunisasi := SasaranImunisasi{DetailImunisasi: make(map[string]DetailImunisasi)}
		for sasaranColumnName := range sasaranColumnMap {
			if sasaranColumnName == USIA_ANAK {
				continue
			}

			cell := sourceColumnMap[sasaranColumnName].Label + strRowIndex

			if IsSourceInvalid(sourceColumnMap, strRowIndex, sourceFile) {
				isRowValid = false
				break
			}

			sasaranCellHandler := map[string]func(){
				NAMA_ANAK:          func() { sasaranImunisasi.NamaAnak = GetCellValue(sourceFile, cell) },
				TANGGAL_LAHIR_ANAK: func() { sasaranImunisasi.TanggalLahirAnak = GetCellValue(sourceFile, cell) },
				JENIS_KELAMIN_ANAK: func() { sasaranImunisasi.JenisKelaminAnak = GetCellValue(sourceFile, cell) },
				NAMA_ORANG_TUA:     func() { sasaranImunisasi.NamaOrangTua = GetCellValue(sourceFile, cell) },
				PUSKESMAS:          func() { sasaranImunisasi.Puskesmas = GetCellValue(sourceFile, cell) },
			}

			if cellHandler, exists := sasaranCellHandler[sasaranColumnName]; exists {
				cellHandler()
			} else {
				imunisasiType := EMPTY_STRING
				for _, imunisasiBayi := range svc.Cfg.ImunisasiBayi {
					if strings.Contains(sasaranColumnName, imunisasiBayi) {
						imunisasiType = imunisasiBayi
						break
					}
				}
				if imunisasiType == EMPTY_STRING {
					for _, imunisasiBaduta := range svc.Cfg.ImunisasiBaduta {
						if strings.Contains(sasaranColumnName, imunisasiBaduta) {
							imunisasiType = imunisasiBaduta
							break
						}
					}
				}

				detailImunisasi := sasaranImunisasi.DetailImunisasi[imunisasiType]

				cellValue := GetCellValue(sourceFile, cell)
				switch {
				case strings.Contains(sasaranColumnName, TANGGAL):
					if _, exists := detailImunisasi.Tanggal[sasaranColumnName]; exists {
						detailImunisasi.Tanggal[sasaranColumnName] = cellValue
					} else {
						tanggal := make(map[string]string)
						tanggal[sasaranColumnName] = cellValue
						detailImunisasi.Tanggal = tanggal
					}
				case strings.Contains(sasaranColumnName, POS):
					if _, exists := detailImunisasi.Pos[sasaranColumnName]; exists {
						detailImunisasi.Pos[sasaranColumnName] = cellValue
					} else {
						pos := make(map[string]string)
						pos[sasaranColumnName] = cellValue
						detailImunisasi.Pos = pos
					}
				case strings.Contains(sasaranColumnName, STATUS):
					if _, exists := detailImunisasi.Status[sasaranColumnName]; exists {
						detailImunisasi.Status[sasaranColumnName] = GetStatusImunisasi(cellValue)
					} else {
						status := make(map[string]int)
						status[sasaranColumnName] = GetStatusImunisasi(cellValue)
						detailImunisasi.Status = status
					}
				}
				sasaranImunisasi.DetailImunisasi[imunisasiType] = detailImunisasi
			}

			if sasaranColumnName == TANGGAL_LAHIR_ANAK {
				sasaranImunisasi.UsiaAnak = sasaranImunisasi.CalculateUsiaAnak()
			}
		}
		rowIndex++

		if isRowValid && sasaranImunisasi.CountNonIdealImmunizations() > 0 {
			sasaranImunisasiList = append(sasaranImunisasiList, sasaranImunisasi)
		}
	}

	SortDataImunisasiAnak(sasaranImunisasiList, func(s SasaranImunisasi) string {
		return s.TanggalLahirAnak
	})

	svc.SasaranImunisasiList = sasaranImunisasiList

	excelFile, err := CreateNewXlsxFile(sourceFile.Ctx, svc)
	if err != nil {
		return nil, err
	}

	kejarType := GetKejarTypeFromContext(sourceFile.Ctx)
	return &XlsxGeneratedFile{
		FileName:     svc.GetFileName(kejarType) + ".xlsx",
		ExcelizeFile: excelFile,
	}, nil
}

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

// GetSasaranColumnMap returns sasaran column map and last column label based on sasaran type
func (svc *SasaranImunisasiService) GetSasaranColumnMap(ctx context.Context) (map[string]Column, string) {
	sasaranType := GetKejarTypeFromContext(ctx)
	if sasaranType == BAYI {
		return svc.SasaranBayiColumnMap, svc.SasaranBayiColumnMap[STATUS_IDL_1].Label
	}
	return svc.SasaranBadutaColumnMap, svc.SasaranBadutaColumnMap[STATUS_IMUNISASI_PCV_3].Label
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

// IsSourceInvalid checks whether the immunization information from source file is valid
// based on the given column and row index. It returns true if invalid, otherwise false.
func IsSourceInvalid(sourceColumnMap map[string]Column, strRowIndex string, sourceFile SourceXlsxFile) bool {
	for name, column := range sourceColumnMap {
		if strings.Contains(name, POS) {
			cellValue := strings.ToLower(GetCellValue(sourceFile, column.Label+strRowIndex))
			if strings.Contains(cellValue, "cibuntu") {
				return true
			}

			if strings.Contains(cellValue, "wanasari") {
				return false
			}

			if strings.Contains(cellValue, "dalam gedung") {
				return false
			}

			if strings.Contains(cellValue, "oleh sistem") {
				return false
			}

			if cellValue == "-" {
				return false
			}
		}
	}
	return false
}

// GetFileName returns title based on kejarType
func (svc *SasaranImunisasiService) GetFileName(sasaranType string) string {
	return "Sasaran Imunisasi " + CapitalizeFirstChar(sasaranType) + SPACE + GetCurrentDateStr()
}

// SetTitle sets the title of the Excel sheet for the generated file
func (svc *SasaranImunisasiService) SetTitle(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.TitleRowAt)
	sasaranType := GetKejarTypeFromContext(newFile.Ctx)
	title := svc.GetFileName(sasaranType)
	sasaranImunisasiMap, lastColumnLabel := svc.GetSasaranColumnMap(newFile.Ctx)

	firstCell := sasaranImunisasiMap[NAMA_ANAK].Label + rowAt
	lastCell := lastColumnLabel + rowAt

	file.SetCellValue(sheetName, firstCell, title)
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
	sasaranKejarMap, _ := svc.GetSasaranColumnMap(newFile.Ctx)
	for _, column := range sasaranKejarMap {
		file.SetColWidth(sheetName, column.Label, column.Label, 32) // TODO: how to autosize?
	}
}
