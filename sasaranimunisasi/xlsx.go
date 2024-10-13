package sasaranimunisasi

import (
	"context"
	"log"

	"github.com/xuri/excelize/v2"
)

// XlsxSourceFile contains information about the source Excel file,
// including its temporary file path, the sheet name, and the opened Excel file.
type XlsxSourceFile struct {
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
	GenerateFile(sourceFile XlsxSourceFile) (*XlsxGeneratedFile, error)
}

// GetCellValue retrieves the value of a cell; returns "-" if an error occurs or the value is empty.
func GetCellValue(sourceFile XlsxSourceFile, cell string) string {
	cellValue, err := sourceFile.ExcelizeFile.GetCellValue(sourceFile.SheetName, cell)
	if err != nil || cellValue == "" {
		return "-"
	}
	return cellValue
}

// Column represents column characteristics of xlsx file
type Column struct {
	Label string
	Width float64
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

// consts for xlsx file
const (
	BLACK_COLOR = "#000000"
	SHEET_NAME  = "Sheet1"
	FONT_TYPE   = "Times New Roman"
)

// CreateNewXlsxFile creates a new Excel file and sets up its styles and structure.
func CreateNewXlsxFile(ctx context.Context, generator NewXlsxGenerator) (*excelize.File, error) {
	excelizeFile := excelize.NewFile()
	index, err := excelizeFile.NewSheet(SHEET_NAME)
	if err != nil {
		return nil, err
	}

	excelizeFile.SetActiveSheet(index)
	excelizeFile.SetDefaultFont(FONT_TYPE)

	newXlsxFile := NewXlsxFile{
		Ctx:            ctx,
		SheetName:      SHEET_NAME,
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

// setStylesForNewFile creates and assigns styles for the title, header, and body.
func setStylesForNewFile(file *excelize.File, newFile *NewXlsxFile) error {
	titleStyle, err := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:      22,
			Bold:      true,
			Color:     BLACK_COLOR,
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
			Color:     BLACK_COLOR,
			VertAlign: "center",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
		Border: []excelize.Border{
			{Type: "left", Style: 1, Color: BLACK_COLOR},
			{Type: "right", Style: 1, Color: BLACK_COLOR},
			{Type: "top", Style: 1, Color: BLACK_COLOR},
			{Type: "bottom", Style: 1, Color: BLACK_COLOR},
		},
	})
	if err != nil {
		log.Printf("Error creating style for header: %v", err)
		return 0
	}

	return style
}
