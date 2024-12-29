namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.TypeConversion;

/// <summary>
/// Specific implementation of <see cref="TypeConverterOptions"/> that adds the Excel cell format to the options that can be used to define the type conversion.
/// Provides options for Excel type conversion, including cell format specification.
/// </summary>
public class ExcelTypeConverterOptions : TypeConverterOptions
{
    /// <summary>
    /// Gets or sets the Excel cell format <see cref="ExcelCellFormats"/>.
    /// </summary>
    public ExcelCellFormats ExcelCellFormat { get; set; }
}