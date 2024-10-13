package sasaranimunisasi

import (
	"strconv"
)

// SasaranImunisasiService manages sasaran imunisasi data and column mapping for bayi and baduta
type SasaranImunisasiService struct {
	Cfg                    *SasaranImunisasiConfig
	SourceFileColumnMap    map[string]Column
	SasaranBayiColumnMap   map[string]Column // represents xlsx column map for the generated file
	SasaranBadutaColumnMap map[string]Column // represents xlsx column map for the generated file
	SasaranImunisasiList   []SasaranImunisasi
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
	SortByStrDate(sasaranImunisasiList, func(s SasaranImunisasi) string {
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

// GetFileName returns title based on sasaranType
func (svc *SasaranImunisasiService) GetFileName(sasaranType string) string {
	return "Sasaran Imunisasi " + CapitalizeFirstChar(sasaranType) + SPACE + GetCurrentDateStr()
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
