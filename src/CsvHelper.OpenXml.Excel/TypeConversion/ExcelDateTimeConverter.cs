namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

/// <summary>
/// Converts Excel date and time values to and from <see cref="DateTime"/> objects.
/// </summary>
public class ExcelDateTimeConverter : DefaultTypeConverter
{
    /// <summary>
    /// Converts the specified string to a <see cref="DateTime"/> object.
    /// </summary>
    /// <param name="text">The string to convert.</param>
    /// <param name="row">The <see cref="IReaderRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A <see cref="DateTime"/> object if the conversion was successful; otherwise, null, or null if <paramref name="text"/> is null or consists only of whitespace.</returns>
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData) => string.IsNullOrWhiteSpace(text) ? null : DateTime.FromOADate(double.Parse(text.Replace('.', ',')));

    /// <summary>
    /// Converts the specified <see cref="DateTime"/> object to a string.
    /// </summary>
    /// <param name="value">The <see cref="DateTime"/> object to convert.</param>
    /// <param name="row">The <see cref="IWriterRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being written.</param>
    /// <returns>A string representation of the <see cref="DateTime"/> object if the conversion was successful; otherwise, null if the <paramref name="value"/> is null.</returns>
    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : ((DateTime)value).ToOADate().ToString().Replace(',', '.');
}