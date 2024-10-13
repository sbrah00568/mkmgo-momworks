package kejar

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

// Create a key for the KejarType value
const kejarTypeKey contextKey = "kejarType"

// getKejarTypeFromContext retrieves kejar type from context
func GetKejarTypeFromContext(ctx context.Context) string {
	if kejarType, ok := ctx.Value(kejarTypeKey).(string); ok {
		return kejarType
	}
	return ""
}

// SasaranKejarService handles the processing of DataImunisasiAnak from Excel
type SasaranKejarService struct {
	DataImunisasiBayiMap   map[string]XlsxColumn
	DataImunisasiBadutaMap map[string]XlsxColumn
	SasaranKejarBayiMap    map[string]XlsxColumn
	SasaranKejarBadutaMap  map[string]XlsxColumn
	ListImunisasiAnak      []*DataImunisasiAnak
}

// NewSasaranKejarService creates a new SasaranKejarService
func NewSasaranKejarService(config *KejarConfig) *SasaranKejarService {
	return &SasaranKejarService{
		DataImunisasiBadutaMap: PopulateXlsxColumnMap(config.ColumnDataImunisasiBaduta),
		SasaranKejarBadutaMap:  PopulateXlsxColumnMap(config.ColumnSasaranKejarBaduta),
		DataImunisasiBayiMap:   PopulateXlsxColumnMap(config.ColumnDataImunisasiBayi),
		SasaranKejarBayiMap:    PopulateXlsxColumnMap(config.ColumnSasaranKejarBayi),
	}
}

// GenerateFile generates an Excel file with the processed DataImunisasiAnak
func (svc *SasaranKejarService) GenerateFile(sourceFile SourceXlsxFile) (*XlsxGeneratedFile, error) {
	var listImunisasiAnak []*DataImunisasiAnak
	rowIndex := 2

	for {
		data := NewDataImunisasiAnak()
		isNoMoreRow, isRowValid := data.PopulateFromRow(svc, rowIndex, sourceFile)
		if isNoMoreRow {
			break
		}

		listImunisasiAnak = data.AppendIfValid(isRowValid, listImunisasiAnak)
		rowIndex++
	}

	SortDataImunisasiAnak(listImunisasiAnak, func(d *DataImunisasiAnak) string {
		return d.TanggalLahirAnak
	})

	svc.ListImunisasiAnak = listImunisasiAnak

	excelFile, err := CreateNewXlsxFile(sourceFile.Ctx, svc)
	if err != nil {
		return nil, err
	}

	kejarType := GetKejarTypeFromContext(sourceFile.Ctx)
	return &XlsxGeneratedFile{
		FileName:     svc.GetKejarTitle(kejarType) + ".xlsx",
		ExcelizeFile: excelFile,
	}, nil
}

// GetColumnMap retrieves columnMap based on kejarType
func (svc *SasaranKejarService) GetColumnMap(kejarType string) map[string]XlsxColumn {
	if kejarType == "bayi" {
		return svc.DataImunisasiBayiMap
	}
	return svc.DataImunisasiBadutaMap
}

// SortDataImunisasiAnak sorts data imunisasi anak based on tanggal lahir (DESC)
func SortDataImunisasiAnak[T any](list []T, dateExtractor func(T) string) {
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

// GetKejarTitle returns title based on kejarType
func (svc *SasaranKejarService) GetKejarTitle(kejarType string) string {
	return "Sasaran Imunisasi Kejar " + CapitalizeFirstChar(kejarType) + " " + GetCurrentDateStr()
}

// SetTitle sets the title of the Excel sheet for the generated file
func (svc *SasaranKejarService) SetTitle(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.TitleRowAt)
	kejarType := GetKejarTypeFromContext(newFile.Ctx)
	title := svc.GetKejarTitle(kejarType)
	sasaranKejarMap := svc.GetSasaranKejarMap(kejarType)

	firstCell := sasaranKejarMap[NamaAnak].Label + rowAt
	var lastCell string
	if kejarType == "bayi" {
		lastCell = sasaranKejarMap[StatusImunisasiIDL1].Label + rowAt
	} else {
		lastCell = sasaranKejarMap[StatusImunisasiPCV3].Label + rowAt
	}

	file.SetCellValue(sheetName, firstCell, title)
	file.MergeCell(sheetName, firstCell, lastCell)
	file.SetCellStyle(sheetName, firstCell, lastCell, newFile.TitleStyle)
}

// SetHeader sets the header row of the Excel sheet
func (svc *SasaranKejarService) SetHeader(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	rowAt := strconv.Itoa(newFile.HeaderRowAt)
	kejarType := GetKejarTypeFromContext(newFile.Ctx)
	sasaranKejarMap := svc.GetSasaranKejarMap(kejarType)

	for _, column := range sasaranKejarMap {
		file.SetCellValue(sheetName, column.Label+rowAt, column.Name)
	}

	firstBodyCell := sasaranKejarMap[NamaAnak].Label + rowAt
	var lastBodyCell string

	if kejarType == "bayi" {
		lastBodyCell = sasaranKejarMap[StatusImunisasiIDL1].Label + rowAt
	} else {
		lastBodyCell = sasaranKejarMap[StatusImunisasiPCV3].Label + rowAt
	}

	file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.HeaderStyle)
}

// GetSasaranKejarMap retrieves sasaranKejarMap based on kejarType
func (svc *SasaranKejarService) GetSasaranKejarMap(kerjaType string) map[string]XlsxColumn {
	if kerjaType == "bayi" {
		return svc.SasaranKejarBayiMap
	}
	return svc.SasaranKejarBadutaMap
}

// SetBody sets the body row of the Excel sheet
func (svc *SasaranKejarService) SetBody(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	kejarType := GetKejarTypeFromContext(newFile.Ctx)
	sasaranKejarMap := svc.GetSasaranKejarMap(kejarType)

	for i, sasaranKejar := range svc.ListImunisasiAnak {
		rowAt := strconv.Itoa(i + newFile.StartBodyRowAt)
		file.SetCellValue(sheetName, sasaranKejarMap[NamaAnak].Label+rowAt, sasaranKejar.NamaAnak)
		file.SetCellValue(sheetName, sasaranKejarMap[UsiaAnak].Label+rowAt, sasaranKejar.UsiaAnak)
		file.SetCellValue(sheetName, sasaranKejarMap[TanggalLahirAnak].Label+rowAt, sasaranKejar.TanggalLahirAnak)
		file.SetCellValue(sheetName, sasaranKejarMap[JenisKelaminAnak].Label+rowAt, sasaranKejar.JenisKelaminAnak)
		file.SetCellValue(sheetName, sasaranKejarMap[NamaOrangTua].Label+rowAt, sasaranKejar.NamaOrangTua)
		file.SetCellValue(sheetName, sasaranKejarMap[Puskesmas].Label+rowAt, sasaranKejar.Puskesmas)
		for key, value := range sasaranKejar.DetailImunisasi {
			file.SetCellValue(sheetName, sasaranKejarMap[TanggalImunisasiPrfx+key].Label+rowAt, value.Tanggal)
			file.SetCellStyle(sheetName, sasaranKejarMap[TanggalImunisasiPrfx+key].Label+rowAt, sasaranKejarMap[TanggalImunisasiPrfx+key].Label+rowAt, newFile.BodyStyle)

			file.SetCellValue(sheetName, sasaranKejarMap[PosImunisasiPrfx+key].Label+rowAt, value.Pos)
			file.SetCellStyle(sheetName, sasaranKejarMap[PosImunisasiPrfx+key].Label+rowAt, sasaranKejarMap[PosImunisasiPrfx+key].Label+rowAt, newFile.BodyStyle)

			file.SetCellValue(sheetName, sasaranKejarMap[StatusImunisasiPrfx+key].Label+rowAt, value.Status)
			file.SetCellStyle(sheetName, sasaranKejarMap[StatusImunisasiPrfx+key].Label+rowAt, sasaranKejarMap[StatusImunisasiPrfx+key].Label+rowAt, newFile.BodyStyle)
		}
	}

	firstBodyCell := sasaranKejarMap[NamaAnak].Label + strconv.Itoa(newFile.StartBodyRowAt)
	lastBodyCell := sasaranKejarMap[Puskesmas].Label + strconv.Itoa(len(svc.ListImunisasiAnak)+newFile.HeaderRowAt)
	file.SetCellStyle(sheetName, firstBodyCell, lastBodyCell, newFile.BodyStyle)
}

// SetColumnWidth sets the column width of the Excel sheet
func (svc *SasaranKejarService) SetColumnWidth(newFile NewXlsxFile) {
	file := newFile.ExcelizeFile
	sheetName := newFile.SheetName
	kejarType := GetKejarTypeFromContext(newFile.Ctx)
	sasaranKejarMap := svc.GetSasaranKejarMap(kejarType)

	for _, column := range sasaranKejarMap {
		file.SetColWidth(sheetName, column.Label, column.Label, float64(column.Length))
	}

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

// CapitalizeFirstChar capitalizes the first character of a string.
func CapitalizeFirstChar(input string) string {
	if len(input) == 0 {
		return input
	}
	return strings.ToUpper(string(input[0])) + input[1:]
}
