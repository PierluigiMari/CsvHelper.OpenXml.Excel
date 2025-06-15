namespace CsvHelper.OpenXml.Excel.TypeConversion;

/// <summary>
/// Enumeration that specifies the types of values that can result from processing an Excel hyperlink.
/// </summary>
public enum ExcelHyperlinkResultantValueTypes : int
{
    /// <summary>
    /// Represents a single string value in the enumeration.
    /// </summary>
    SingleString = 0,

    /// <summary>
    /// Represents a mode where a single URI is used.
    /// </summary>
    SingleUri = 1,

    /// <summary>
    /// Represents a <see cref="ValueTuple{T1, T2}" /> where T1 and T2 are respectively of type <see cref="string"/>, <see cref="string"/>.
    /// </summary>
    TupleStringString = 2,

    /// <summary>
    /// Represents a <see cref="ValueTuple{T1, T2}" /> where T1 and T2 are respectively of type <see cref="string"/>, <see cref="Uri"/>.
    /// </summary>
    TupleStringUri = 3
}