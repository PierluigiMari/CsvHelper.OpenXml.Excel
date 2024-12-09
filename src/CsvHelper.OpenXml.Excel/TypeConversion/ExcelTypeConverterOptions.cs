namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.TypeConversion;

public class ExcelTypeConverterOptions : TypeConverterOptions
{
    public ExcelCellFormats ExcelCellFormat { get; set; }
}