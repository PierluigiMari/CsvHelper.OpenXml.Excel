namespace CsvHelper.OpenXml.Excel.TypeConversion.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using Shouldly;
using System;
using Xunit;


public class ExcelTimeOnlyConverterTests
{
    #region Fields

    private readonly ExcelTimeOnlyConverter ExcelTimeOnlyConverter = new ExcelTimeOnlyConverter();
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
        object? ConversionResult = ExcelTimeOnlyConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData("0.5", 12, 0, 0)] // 12:00:00
    [InlineData("0,75", 18, 0, 0)] // 18:00:00 (comma as decimal separator)
    [InlineData("0.0", 0, 0, 0)] // 00:00:00
    [InlineData("0.999988425925926", 23, 59, 59)] // 23:59:59
    public void ConvertFromString_ShouldConvertOADateStringToTimeOnlyCorrectly(string input, int expectedhour, int expectedminute, int expectedsecond)
    {
        object? ConversionResult = ExcelTimeOnlyConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeOfType<TimeOnly>();
        ((TimeOnly)ConversionResult!).ShouldBe(new TimeOnly(expectedhour, expectedminute, expectedsecond));
    }

    [Theory]
    [InlineData("notanumber")]
    [InlineData("45239-foo")]
    public void ConvertFromString_ShouldThrowsFormatExceptionForInvalidInput(string input)
    {
        Should.Throw<FormatException>(() => ExcelTimeOnlyConverter.ConvertFromString(input, NullReaderRow, MemberMapData));
    }

    [Fact]
    public void ConvertToString_ShouldReturnsNullWhenIsNull()
    {
        string? ConversionResult = ExcelTimeOnlyConverter.ConvertToString(null, NullWriterRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData(0, 0, 0, "0")]
    [InlineData(12, 0, 0, "0.5")]
    [InlineData(18, 0, 0, "0.75")]
    [InlineData(23, 59, 59, "0.999988425925926")]
    public void ConvertToString_ShouldConvertsTimeOnlyToOADateStringCorrectly(int hour, int minute, int second, string expected)
    {
        TimeOnly Time = new TimeOnly(hour, minute, second);
        string?  ConversionResult = ExcelTimeOnlyConverter.ConvertToString(Time, NullWriterRow, MemberMapData);

        ConversionResult.ShouldStartWith(expected.Replace(',', '.'));
    }

    [Fact]
    public void ConvertToString_ShouldThrowsInvalidCastExceptionForNonTimeOnly()
    {
        Should.Throw<InvalidCastException>(() => ExcelTimeOnlyConverter.ConvertToString(123, NullWriterRow, MemberMapData));
    }

    #endregion
}