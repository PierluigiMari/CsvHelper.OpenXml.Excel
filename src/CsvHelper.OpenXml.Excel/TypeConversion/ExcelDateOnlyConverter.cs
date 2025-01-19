namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

/// <summary>
/// Converts Excel date values to and from <see cref="DateOnly"/> objects.
/// </summary>
public class ExcelDateOnlyConverter : DefaultTypeConverter
{
    /// <summary>
    /// Converts the specified string to a <see cref="DateOnly"/> object.
    /// </summary>
    /// <param name="text">The string to convert.</param>
    /// <param name="row">The <see cref="IReaderRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A <see cref="DateOnly"/> object if the conversion was successful; otherwise, null.</returns>
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData) => string.IsNullOrWhiteSpace(text) ? null : DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(text.Replace('.', ','))));

    /// <summary>
    /// Converts the specified <see cref="DateOnly"/> object to a string.
    /// </summary>
    /// <param name="value">The <see cref="DateOnly"/> object to convert.</param>
    /// <param name="row">The <see cref="IWriterRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A string representation of the <see cref="DateOnly"/> object if the conversion was successful; otherwise, null.</returns>
    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : ((DateOnly)value).ToDateTime(new TimeOnly()).ToOADate().ToString().Replace(',', '.');
}