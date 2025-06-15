namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

/// <summary>
/// Converts an Excel date time string values or an Excel date time values to and from <see cref="DateTimeOffset"/> objects.
/// </summary>
/// <param name="datetimekind">The kind of date and time (UTC, Local, or Unspecified).</param>
/// <param name="offsetfromutc">The offset from UTC.</param>
public class ExcelDateTimeOffsetTextConverter(DateTimeKind datetimekind = DateTimeKind.Utc, TimeSpan? offsetfromutc = null) : DefaultTypeConverter
{
    private readonly TimeSpan OffsetFromUtc = offsetfromutc ?? datetimekind switch
    {
        DateTimeKind.Utc => TimeSpan.Zero,
        DateTimeKind.Local => TimeZoneInfo.Local.GetUtcOffset(DateTime.UtcNow),
        DateTimeKind.Unspecified => TimeSpan.Zero,
        _ => throw new ArgumentOutOfRangeException(nameof(datetimekind), datetimekind, "Invalid DateTimeKind value.")
    };

    private readonly DateTimeKind DateTimeKind = datetimekind;

    /// <summary>
    /// Converts the specified string to a <see cref="DateTimeOffset"/>.
    /// </summary>
    /// <param name="text">The string to convert.</param>
    /// <param name="row">The <see cref="IReaderRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>A <see cref="DateTimeOffset"/> object if the conversion succeeded; otherwise, null, or null if <paramref name="text"/> is null or consists only of whitespace.</returns>
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        const string DelimiterMinus = "-";

        int DelimiterMinusIndex = text.LastIndexOf(DelimiterMinus, StringComparison.Ordinal);

        if (text.Contains('+') || DelimiterMinusIndex > 19 || text.Contains('Z'))
        {
            if (DateTimeOffset.TryParse(text, out DateTimeOffset datetimeoffset))
                return datetimeoffset;
            else
                return null;
        }

        TimeSpan OffsetFromUtcForLocal = TimeZoneInfo.Local.GetUtcOffset(DateTime.Parse(text));

        if (DateTimeKind == DateTimeKind.Local && OffsetFromUtcForLocal != OffsetFromUtc)
            return DateTime.TryParse(text, out DateTime datetime) ? new DateTimeOffset(DateTime.SpecifyKind(datetime, DateTimeKind), OffsetFromUtcForLocal) : null;
        else
            return DateTime.TryParse(text, out DateTime datetime) ? new DateTimeOffset(DateTime.SpecifyKind(datetime, DateTimeKind), OffsetFromUtc) : null;
    }

    /// <summary>
    /// Converts the specified <see cref="DateTimeOffset"/> object to a string.
    /// </summary>
    /// <param name="value">The <see cref="DateTimeOffset"/> object to convert.</param>
    /// <param name="row">The <see cref="IWriterRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being written.</param>
    /// <returns>A string representation of the <see cref="DateTimeOffset"/> object if the conversion was successful; otherwise, null if the <paramref name="value"/> is null.</returns>
    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : ((DateTimeOffset)value).ToString(memberMapData.TypeConverterOptions.CultureInfo);
}