namespace CsvHelper.OpenXml.Excel.TypeConversion;

/// <summary>
/// Enumeration that specifies the various cell formats that can be applied to Excel cell headers.
/// </summary>
public enum ExcelCellHeaderFormats : int
{
    /// <summary>
    /// Default cell format.
    /// </summary>
    Default = 0,
    /// <summary>
    /// Default cell format with bold text.
    /// </summary>
    DefaultBold = 1,
    /// <summary>
    /// Default cell format with bold and centered text.
    /// </summary>
    DefaultBoldCentered = 2,
}