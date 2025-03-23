namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

/// <summary>
/// Converts an Excel date time values to and from <see cref="DateTimeOffset"/> objects.
/// </summary>
/// <remarks>Excel date and time values are treated as UTC (Coordinated Universal Time) values.</remarks>
public class ExcelDateTimeOffsetConverter(DateTimeKind datetimekind = DateTimeKind.Utc) : DefaultTypeConverter
{
    private readonly DateTimeKind DateTimeKind = datetimekind;

    /// <summary>
    /// Converts the specified string to a <see cref="DateTimeOffset"/> object.
    /// </summary>
    /// <param name="text">The string to convert.</param>
    /// <param name="row">The <see cref="IReaderRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A <see cref="DateTimeOffset"/> object if the conversion succeeded; otherwise, null.</returns>
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData) => string.IsNullOrWhiteSpace(text) ? null : new DateTimeOffset(DateTime.SpecifyKind(DateTime.FromOADate(double.Parse(text.Replace('.', ','))), DateTimeKind));

    /// <summary>
    /// Converts the specified <see cref="DateTimeOffset"/> object to a string.
    /// </summary>
    /// <param name="value">The <see cref="DateTimeOffset"/> object to convert.</param>
    /// <param name="row">The <see cref="IWriterRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A string representation of the <see cref="DateTimeOffset"/> object if the conversion was successful; otherwise, null.</returns>
    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : DateTimeKind switch
    {
        DateTimeKind.Utc => ((DateTimeOffset)value).UtcDateTime.ToOADate().ToString().Replace(',', '.'),
        DateTimeKind.Local => ((DateTimeOffset)value).LocalDateTime.ToOADate().ToString().Replace(',', '.'),
        DateTimeKind.Unspecified => ((DateTimeOffset)value).DateTime.ToOADate().ToString().Replace(',', '.'),
        _ => throw new ArgumentOutOfRangeException(nameof(DateTimeKind), DateTimeKind, "Invalid DateTimeKind value.")
    };
}