namespace CsvHelper.OpenXml.Excel.TypeConversion;

/// <summary>
/// Enumeration to categorize and handle hyperlinks appropriately when working with Excel-related functionality.
/// </summary>
public enum ExcelHyperlinkTypes : int
{
    /// <summary>
    /// Default hyperlink type.
    /// </summary>
    Default = 0,
    /// <summary>
    /// Hyperlink type for web URLs.
    /// </summary>
    WebUrl = 1,
    /// <summary>
    /// Hyperlink type for email addresses.
    /// </summary>
    Email = 2,
    /// <summary>
    /// Hyperlink type for file paths.
    /// </summary>
    FilePath = 3,
    /// <summary>
    /// Hyperlink type for internal document links.
    /// </summary>
    InternalLink = 4
}