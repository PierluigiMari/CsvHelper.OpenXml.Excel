namespace CsvHelper.OpenXml.Excel.TypeConversion.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using Shouldly;
using Xunit;

public class ExcelDateTimeOffsetConverterTests
{
    #region Fields

    private readonly MemberMapData MemberMapData = new MemberMapData(null);
    private readonly IReaderRow NullReaderRow = null!;
    private readonly IWriterRow NullWriterRow = null!;

    #endregion

    #region Test Methods

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ConvertFromString_ShouldReturnsNullWhenIsNullOrWhitespace(string? input)
    {
        ExcelDateTimeOffsetConverter ExcelDateTimeOffsetConverter = new ExcelDateTimeOffsetConverter();

        object? ConversionResult = ExcelDateTimeOffsetConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData("45292", DateTimeKind.Utc, null, 2024, 1, 1, 0, 0, 0, 0)]
    [InlineData("45292,5", DateTimeKind.Utc, null, 2024, 1, 1, 12, 0, 0, 0)]
    [InlineData("45292.5", DateTimeKind.Utc, null, 2024, 1, 1, 12, 0, 0, 0)]
    [InlineData("45292", DateTimeKind.Local, "01:00:00", 2024, 1, 1, 0, 0, 0, 1)]
    [InlineData("45292", DateTimeKind.Local, null, 2024, 1, 1, 0, 0, 0, 1)]
    [InlineData("45292", DateTimeKind.Unspecified, null, 2024, 1, 1, 0, 0, 0, 0)]
    [InlineData("45292", DateTimeKind.Unspecified, "02:00:00", 2024, 1, 1, 0, 0, 0, 2)]
    public void ConvertFromString_ShouldConvertOADateStringToDateTimeOffsetCorrectly(string input, DateTimeKind kind, string? offsetstring, int expectedyear, int expectedmonth, int expectedday, int expectedhour, int expectedminute, int expectedsecond, int expectedoffset)
    {
        TimeSpan? offset = offsetstring != null ? TimeSpan.Parse(offsetstring) : null;

        ExcelDateTimeOffsetConverter ExcelDateTimeOffsetConverter = new ExcelDateTimeOffsetConverter(kind, offset);

        object? ConversionResult = ExcelDateTimeOffsetConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeOfType<DateTimeOffset>();
        ((DateTimeOffset)ConversionResult!).ShouldBe(new DateTimeOffset(expectedyear, expectedmonth, expectedday, expectedhour, expectedminute, expectedsecond, 0, TimeSpan.FromHours(expectedoffset)));
    }

    [Fact]
    public void ConvertToString_ShouldReturnsNullWhenIsNull()
    {
        ExcelDateTimeOffsetConverter ExcelDateTimeOffsetConverter = new ExcelDateTimeOffsetConverter(DateTimeKind.Utc);

        object? ConversionResult = ExcelDateTimeOffsetConverter.ConvertToString(null, NullWriterRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData(2024, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc, "45292")]
    [InlineData(2024, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified, "45292")]
    [InlineData(2024, 1, 1, 12, 0, 0, 0, DateTimeKind.Utc, "45292.5")]
    public void ConvertToString_ShouldConvertsDateTimeOffsetToOADateStringCorrectly(int year, int month, int day, int hour, int minute, int second, int offsethours, DateTimeKind kind, string expectedoadate)
    {
        TimeSpan offset = TimeSpan.FromHours(offsethours);

        ExcelDateTimeOffsetConverter ExcelDateTimeOffsetConverter = new ExcelDateTimeOffsetConverter(kind);

        string? ConversionResult = ExcelDateTimeOffsetConverter.ConvertToString(new DateTimeOffset(year, month, day, hour, minute, second, offset), NullWriterRow, MemberMapData);

        ConversionResult.ShouldStartWith(expectedoadate);
    }

    #endregion
}