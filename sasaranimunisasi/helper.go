package sasaranimunisasi

import (
	"context"
	"fmt"
	"log"
	"sort"
	"strings"
	"time"
)

// SasaranImunisasiConfig holds apps configuration for sasaran imunisasi
type SasaranImunisasiConfig struct {
	ColumnName             []string `yaml:"column_name"`
	DetailImunisasi        []string `yaml:"detail_imunisasi"`
	DetailImunisasiLengkap []string `yaml:"detail_imunisasi_lengkap"`
	ImunisasiBayi          []string `yaml:"imunisasi_bayi"`
	ImunisasiBaduta        []string `yaml:"imunisasi_baduta"`
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

// SortByStrDate sorts a list of generic items based on a string-formatted date extracted by the dateExtractor function.
// The date format expected is "YYYY-MM-DD" (e.g., "2024-01-01"). If a date cannot be parsed, it logs an error and
// keeps the current order for the erroneous item.
func SortByStrDate[T any](list []T, dateExtractor func(T) string) {
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
