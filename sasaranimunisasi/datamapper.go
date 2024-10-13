package sasaranimunisasi

import (
	"context"
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"
)

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
