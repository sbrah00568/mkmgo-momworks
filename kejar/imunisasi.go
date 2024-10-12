package kejar

import (
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"
)

// consts for sasaran imunisasi kejar
const (
	NamaAnak             = "nama_anak"
	UsiaAnak             = "usia_anak"
	TanggalLahirAnak     = "tanggal_lahir_anak"
	JenisKelaminAnak     = "jenis_kelamin_anak"
	NamaOrangTua         = "nama_orang_tua"
	Puskesmas            = "puskesmas"
	StatusImunisasiPCV3  = "status_imunisasi_pcv_3"
	StatusImunisasiIDL1  = "status_imunisasi_idl_1"
	Wanasari             = "WANASARI"
	StatusImunisasiPrfx  = "status_imunisasi_"
	TanggalImunisasiPrfx = "tanggal_imunisasi_"
	PosImunisasiPrfx     = "pos_imunisasi_"
	EmptyString          = ""
)

// DataImunisasiAnak represents general immunization data for both bayi and baduta
type DataImunisasiAnak struct {
	NamaAnak         string                     `json:"namaAnak"`
	UsiaAnak         string                     `json:"usiaAnak"`
	TanggalLahirAnak string                     `json:"tanggalLahirAnak"`
	JenisKelaminAnak string                     `json:"jenisKelaminAnak"`
	NamaOrangTua     string                     `json:"namaOrangTua"`
	Puskesmas        string                     `json:"puskesmas"`
	MapImunisasi     map[string]DetailImunisasi `json:"mapImunisasi"`
}

// DetailImunisasi represents detail data of given immunization
type DetailImunisasi struct {
	TanggalImunisasi string
	PosImunisasi     string
	Status           int
}

// NewDataImunisasiAnak creates a new DataImunisasiAnak instance
func NewDataImunisasiAnak() *DataImunisasiAnak {
	return &DataImunisasiAnak{
		MapImunisasi: make(map[string]DetailImunisasi),
	}
}

// SetCellValue sets the appropriate cell value for DataImunisasiAnak
func (data *DataImunisasiAnak) SetCellValue(columnCode, cell string, sourceFile SourceXlsxFile) {
	cellHandlers := map[string]func(){
		NamaAnak:         func() { data.NamaAnak = GetCellValue(sourceFile, cell) },
		TanggalLahirAnak: func() { data.TanggalLahirAnak = GetCellValue(sourceFile, cell) },
		JenisKelaminAnak: func() { data.JenisKelaminAnak = GetCellValue(sourceFile, cell) },
		NamaOrangTua:     func() { data.NamaOrangTua = GetCellValue(sourceFile, cell) },
	}

	if handler, exists := cellHandlers[columnCode]; exists {
		handler()
	} else if imunisasiCode := GetMapImunisasiCode(columnCode); imunisasiCode != EmptyString {
		imunisasi := data.MapImunisasi[imunisasiCode]
		imunisasi.SetDetailImunisasi(columnCode, cell, sourceFile)
		data.MapImunisasi[imunisasiCode] = imunisasi
	}
}

// SetDetailImunisasi sets detail imunisasi for DataImunisasiAnak
func (imunisasi *DetailImunisasi) SetDetailImunisasi(columnCode, cell string, sourceFile SourceXlsxFile) {
	switch {
	case strings.Contains(columnCode, StatusImunisasiPrfx):
		imunisasi.Status = GetStatusImunisasi(GetCellValue(sourceFile, cell))
	case strings.Contains(columnCode, TanggalImunisasiPrfx):
		imunisasi.TanggalImunisasi = GetCellValue(sourceFile, cell)
	case strings.Contains(columnCode, PosImunisasiPrfx):
		imunisasi.PosImunisasi = GetCellValue(sourceFile, cell)
	}
}

// GetMapImunisasiCode returns map imunisasi code based on prefix when found, else returns empty
func GetMapImunisasiCode(columnCode string) string {
	prefixDetailImunisasi := map[string]string{
		StatusImunisasiPrfx:  EmptyString,
		TanggalImunisasiPrfx: EmptyString,
		PosImunisasiPrfx:     EmptyString,
	}

	for key, value := range prefixDetailImunisasi {
		if strings.Contains(columnCode, key) {
			return strings.ReplaceAll(columnCode, key, value)
		}
	}
	return EmptyString
}

// GetStatusImunisasi returns an integer status based on the input string.
// "ideal" corresponds to 0, while any other status corresponds to 1.
func GetStatusImunisasi(status string) int {
	if status == "ideal" {
		return 0
	}
	return 1
}

// CountNonIdealImmunizations counts immunizations that are not zero (non-ideal)
func (data *DataImunisasiAnak) CountNonIdealImmunizations() int {
	count := 0
	for _, value := range data.MapImunisasi {
		if value.Status != 0 {
			count++
		}
	}
	return count
}

// PopulateFromRow populates DataImunisasiAnak from a given excel file
func (data *DataImunisasiAnak) PopulateFromRow(svc *SasaranKejarService, rowIndex int, sourceFile SourceXlsxFile) (bool, bool) {
	isNoMoreRow, isRowValid := false, false

	columnMap := svc.GetColumnMap(GetKejarTypeFromContext(sourceFile.Ctx))
	for _, column := range columnMap {
		if CheckForEndOfFile(column, rowIndex, sourceFile) {
			isNoMoreRow = true
			break
		}

		if isRowValid = IsImunisasiValid(column, rowIndex, sourceFile); isRowValid {
			data.SetCellValue(column.Code, column.Label+strconv.Itoa(rowIndex), sourceFile)
		} else {
			break
		}
	}

	return isNoMoreRow, isRowValid
}

// IsImunisasiValid checks whether the immunization information is valid based on the given column and row index.
// It returns true if valid, otherwise false.
func IsImunisasiValid(column XlsxColumn, rowIndex int, sourceFile SourceXlsxFile) bool {
	if !strings.Contains(column.Code, "_pemberi_imunisasi_") && !strings.Contains(column.Code, "pos_imunisasi_") {
		return true
	}

	cellValue := strings.ToLower(GetCellValue(sourceFile, column.Label+strconv.Itoa(rowIndex)))

	validValues := []string{"wanasari", "dalam gedung", "oleh sistem"}
	for _, valid := range validValues {
		if strings.Contains(cellValue, valid) || cellValue == "-" {
			return true
		}
	}

	return false
}

// AppendIfValid adds an data imunisasi anak to the list if valid
func (data *DataImunisasiAnak) AppendIfValid(isValid bool, list []*DataImunisasiAnak) []*DataImunisasiAnak {
	if data.CountNonIdealImmunizations() > 0 && isValid {
		data.Puskesmas = Wanasari
		data.UsiaAnak = data.CalculateUsiaAnak()
		list = append(list, data)
	}
	return list
}

// CalculateUsiaAnak calculates the age of a child based on their birth date in the format "YYYY-MM-DD".
// It returns a string indicating the age in months and days.
func (data *DataImunisasiAnak) CalculateUsiaAnak() string {
	birthDate, err := time.Parse("2006-01-02", data.TanggalLahirAnak)
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
