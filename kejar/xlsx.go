package kejar

import (
	"context"
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// XlsxColumn represents a configuration for a column in an Excel sheet, including
// its code, index, name, label, and length.
type XlsxColumn struct {
	Code   string `yaml:"code"`   // Column code (0, 1, 2, 3, 4, etc.)
	Index  int    `yaml:"index"`  // Column index
	Name   string `yaml:"name"`   // Column name
	Label  string `yaml:"label"`  // Column label (A, B, C, D, etc.)
	Length int    `yaml:"length"` // Max length of the column data
}

// SourceXlsxFile contains information about the source Excel file,
// including its temporary file path, the sheet name, and the opened Excel file.
type SourceXlsxFile struct {
	Ctx          context.Context
	TempFilePath string
	SheetName    string
	ExcelizeFile *excelize.File
}

// XlsxGeneratedFile holds the generated Excel file details,
// including its filename and the Excelize file pointer.
type XlsxGeneratedFile struct {
	FileName     string
	ExcelizeFile *excelize.File
}

// XlsxFileTransformer is an interface defining the method to generate a new Excel file
// from a source Excel file.
type XlsxFileTransformer interface {
	GenerateFile(sourceFile SourceXlsxFile) (*XlsxGeneratedFile, error)
}

// CheckForEndOfFile checks if the end of the file is reached for a given column and row.
func CheckForEndOfFile(column XlsxColumn, rowIndex int, sourceFile SourceXlsxFile) bool {
	return column.Code == "id" && GetCellValue(sourceFile, column.Label+strconv.Itoa(rowIndex)) == "-"
}

// PopulateXlsxColumnMap creates a map of XlsxColumn indexed by their code.
func PopulateXlsxColumnMap(columns []XlsxColumn) map[string]XlsxColumn {
	columnMap := make(map[string]XlsxColumn, len(columns))
	for _, col := range columns {
		columnMap[col.Code] = col
	}
	return columnMap
}

// GetCellValue retrieves the value of a cell; returns "-" if an error occurs or the value is empty.
func GetCellValue(sourceFile SourceXlsxFile, cell string) string {
	cellValue, err := sourceFile.ExcelizeFile.GetCellValue(sourceFile.SheetName, cell)
	if err != nil || cellValue == "" {
		return "-"
	}
	return cellValue
}

// NewXlsxFile represents the structure of a new Excel file.
type NewXlsxFile struct {
	Ctx            context.Context
	SheetName      string
	ExcelizeFile   *excelize.File
	TitleRowAt     int
	HeaderRowAt    int
	StartBodyRowAt int
	TitleStyle     int
	HeaderStyle    int
	BodyStyle      int
}

// NewXlsxGenerator defines methods for generating a new Excel file.
type NewXlsxGenerator interface {
	SetTitle(newFile NewXlsxFile)
	SetHeader(newFile NewXlsxFile)
	SetBody(newFile NewXlsxFile)
	SetColumnWidth(newFile NewXlsxFile)
}

// CreateNewXlsxFile creates a new Excel file and sets up its styles and structure.
func CreateNewXlsxFile(ctx context.Context, generator NewXlsxGenerator) (*excelize.File, error) {
	excelizeFile := excelize.NewFile()
	sheetName := "Sheet1"
	index, err := excelizeFile.NewSheet(sheetName)
	if err != nil {
		return nil, err
	}

	excelizeFile.SetActiveSheet(index)
	excelizeFile.SetDefaultFont("Times New Roman")

	newXlsxFile := NewXlsxFile{
		Ctx:            ctx,
		SheetName:      sheetName,
		ExcelizeFile:   excelizeFile,
		TitleRowAt:     1,
		HeaderRowAt:    3,
		StartBodyRowAt: 4,
	}

	if err := setStylesForNewFile(excelizeFile, &newXlsxFile); err != nil {
		return nil, err
	}

	generator.SetTitle(newXlsxFile)
	generator.SetHeader(newXlsxFile)
	generator.SetBody(newXlsxFile)
	generator.SetColumnWidth(newXlsxFile)

	return excelizeFile, nil
}

const blackColor = "#000000"

// setStylesForNewFile creates and assigns styles for the title, header, and body.
func setStylesForNewFile(file *excelize.File, newFile *NewXlsxFile) error {
	titleStyle, err := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:      22,
			Bold:      true,
			Color:     blackColor,
			VertAlign: "center",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		newFile.TitleStyle = 0 // Default value if an error occurs
	} else {
		newFile.TitleStyle = titleStyle
	}

	newFile.HeaderStyle = SetXlsxStyle(file, true)
	newFile.BodyStyle = SetXlsxStyle(file, false)

	return nil
}

// SetXlsxStyle creates a style for Excel cells based on whether it is a header.
func SetXlsxStyle(file *excelize.File, isHeader bool) int {
	style, err := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:      12,
			Bold:      isHeader,
			Color:     blackColor,
			VertAlign: "center",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
		Border: []excelize.Border{
			{Type: "left", Style: 1, Color: blackColor},
			{Type: "right", Style: 1, Color: blackColor},
			{Type: "top", Style: 1, Color: blackColor},
			{Type: "bottom", Style: 1, Color: blackColor},
		},
	})
	if err != nil {
		log.Printf("Error creating style for header: %v", err)
		return 0
	}

	return style
}
