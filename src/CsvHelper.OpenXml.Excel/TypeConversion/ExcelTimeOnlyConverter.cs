namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

/// <summary>
/// Converts Excel time values to and from <see cref="TimeOnly"/> objects.
/// </summary>
public class ExcelTimeOnlyConverter : DefaultTypeConverter
{
    /// <summary>
    /// Converts the specified string to a <see cref="TimeOnly"/> object.
    /// </summary>
    /// <param name="text">The string to convert.</param>
    /// <param name="row">The <see cref="IReaderRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A <see cref="TimeOnly"/> object if the conversion was successful; otherwise, null.</returns>
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData) => text is null ? null : TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(text.Replace('.', ','))));

    /// <summary>
    /// Converts the specified <see cref="TimeOnly"/> object to a string.
    /// </summary>
    /// <param name="value">The <see cref="TimeOnly"/> object to convert.</param>
    /// <param name="row">The <see cref="IWriterRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A string representation of the <see cref="TimeOnly"/> object if the conversion was successful; otherwise, null.</returns>
    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : new DateOnly(1899, 12, 30).ToDateTime((TimeOnly)value).ToOADate().ToString().Replace(',', '.');
}