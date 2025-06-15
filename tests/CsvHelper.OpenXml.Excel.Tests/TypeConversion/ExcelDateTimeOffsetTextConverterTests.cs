namespace CsvHelper.OpenXml.Excel.TypeConversion.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using CsvHelper.TypeConversion;
using FakeItEasy;
using Shouldly;
using System.Globalization;
using Xunit;

public class ExcelDateTimeOffsetTextConverterTests
{
    #region Fields

    private readonly MemberMapData MemberMapData = A.Fake<MemberMapData>();
    private readonly IReaderRow NullReaderRow = null!;
    private readonly IWriterRow NullWriterRow = null!;

    #endregion

    #region Constructors

    public ExcelDateTimeOffsetTextConverterTests()
    {
        TypeConverterOptions ConverterOptions = new TypeConverterOptions { CultureInfo = CultureInfo.InvariantCulture };
        A.CallTo(() => MemberMapData.TypeConverterOptions).Returns(ConverterOptions);
    }

    #endregion

    #region Test Methods

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ConvertFromString_ShouldReturnsNullWhenIsNullOrWhiteSpace(string? input)
    {
        ExcelDateTimeOffsetTextConverter ExcelDateTimeOffsetTextConverter = new ExcelDateTimeOffsetTextConverter();

        object? ConversionResult = ExcelDateTimeOffsetTextConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData("2024-01-01T12:34:56+99:00")]
    public void ConvertFromString_ShouldReturnsNullWhenInvalidOffsetString(string input)
    {
        ExcelDateTimeOffsetTextConverter ExcelDateTimeOffsetTextConverter = new ExcelDateTimeOffsetTextConverter();

        object? ConversionResult = ExcelDateTimeOffsetTextConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData("2024-01-01T12:34:56 +02:00", 2024, 1, 1, 12, 34, 56, 2)]
    [InlineData("2024-01-01T12:34:56 -05:00", 2024, 1, 1, 12, 34, 56, -5)]
    [InlineData("2024-01-01 12:34:56 +00:00", 2024, 1, 1, 12, 34, 56, 0)]
    [InlineData("2024-01-01T00:00:00Z", 2024, 1, 1, 0, 0, 0, 0)]
    public void ConvertFromString_ShouldConvertOADateStringToDateTimeOffsetCorrectly(string input, int expectedyear, int expectedmonth, int expectedday, int expectedhour, int expectedminute, int expectedsecond, int expectedoffset)
    {
        ExcelDateTimeOffsetTextConverter ExcelDateTimeOffsetTextConverter = new ExcelDateTimeOffsetTextConverter();

        object? ConversionResult = ExcelDateTimeOffsetTextConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeOfType<DateTimeOffset>();
        ((DateTimeOffset)ConversionResult!).ShouldBe(new DateTimeOffset(expectedyear, expectedmonth, expectedday, expectedhour, expectedminute, expectedsecond, 0, TimeSpan.FromHours(expectedoffset)));
    }

    [Theory]
    [InlineData("2024-01-01 12:34:56", DateTimeKind.Utc, 0)]
    [InlineData("2024-01-01 12:34:56", DateTimeKind.Local, 1)]
    [InlineData("2024-01-01 12:34:56", DateTimeKind.Unspecified, 0)]
    public void ConvertFromString_ShouldConvertOADateStringToDateTimeOffsetWithoutOffsetUsingKindAndOffsetCorrectly(string input, DateTimeKind kind, int expectedoffsethours)
    {
        ExcelDateTimeOffsetTextConverter ExcelDateTimeOffsetTextConverter = new ExcelDateTimeOffsetTextConverter(kind);

        object? ConversionResult = ExcelDateTimeOffsetTextConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeOfType<DateTimeOffset>();
        ((DateTimeOffset)ConversionResult!).ShouldBe(new DateTimeOffset(2024, 1, 1, 12, 34, 56, 0, TimeSpan.FromHours(expectedoffsethours)));
    }

    [Fact]
    public void ConvertToString_ShouldReturnNullWhenIsNull()
    {
        ExcelDateTimeOffsetTextConverter ExcelDateTimeOffsetTextConverter = new ExcelDateTimeOffsetTextConverter();

        string? ConversionResult = ExcelDateTimeOffsetTextConverter.ConvertToString(null, NullWriterRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData(2024, 1, 1, 12, 34, 56, 2, "01/01/2024 12:34:56 +02:00")]
    [InlineData(2024, 1, 1, 0, 0, 0, 0, "01/01/2024 00:00:00 +00:00")]
    public void ConvertToString_ShouldConvertsDateTimeOffsetToOADateStringCorrectly(int year, int month, int day, int hour, int minute, int second, int offsetHours, string expected)
    {
        ExcelDateTimeOffsetTextConverter ExcelDateTimeOffsetTextConverter = new ExcelDateTimeOffsetTextConverter();

        string? ConversionResult = ExcelDateTimeOffsetTextConverter.ConvertToString(new DateTimeOffset(year, month, day, hour, minute, second, TimeSpan.FromHours(offsetHours)), NullWriterRow, MemberMapData);

        ConversionResult.ShouldBe(expected);
    }

    #endregion
}