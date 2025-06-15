namespace CsvHelper.OpenXml.Excel.TypeConversion.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using Shouldly;
using System;
using Xunit;

public class ExcelDateTimeConverterTests
{
    #region Fields

    private readonly ExcelDateTimeConverter ExcelDateOnlyConverter = new ExcelDateTimeConverter();
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
        object? ConversionResult = ExcelDateOnlyConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [InlineData("45237", 2023, 11, 7, 0, 0, 0)]
    [InlineData("45239.5", 2023, 11, 9, 12, 0, 0)]
    [InlineData("45239,25", 2023, 11, 9, 6, 0, 0)]
    public void ConvertFromString_ShouldConvertOADateStringToDateTimeCorrectly(string input, int expectedyear, int expectedmonth, int expectedday, int expectedhour, int expectedminute, int expectedsecond)
    {
        object? ConversionResult = ExcelDateOnlyConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeOfType<DateTime>();
        ((DateTime)ConversionResult!).ShouldBe(new DateTime(expectedyear, expectedmonth, expectedday, expectedhour, expectedminute, expectedsecond));
    }

    [Theory]
    [InlineData("notanumber")]
    [InlineData("45239-foo")]
    public void ConvertFromString_ShouldThrowsFormatExceptionForInvalidInput(string input)
    {
        Should.Throw<FormatException>(() => ExcelDateOnlyConverter.ConvertFromString(input, NullReaderRow, MemberMapData));
    }

    [Fact]
    public void ConvertToString_ShouldReturnsNullWhenIsNull()
    {
        string? ConversionResult = ExcelDateOnlyConverter.ConvertToString(null, NullWriterRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Fact]
    public void ConvertToString_ShouldConvertsDateTimeToOADateStringCorrectly()
    {
        DateTime Date = new DateTime(2023, 11, 7, 15, 30, 0);
        string ExpectedOADate = Date.ToOADate().ToString().Replace(',', '.');

        string? ConversionResult = ExcelDateOnlyConverter.ConvertToString(Date, NullWriterRow, MemberMapData);
        ConversionResult.ShouldBe(ExpectedOADate);
    }

    [Fact]
    public void ConvertToString_ShouldThrowsInvalidCastExceptionForNonDateTime()
    {
        Should.Throw<InvalidCastException>(() => ExcelDateOnlyConverter.ConvertToString("notadate", NullWriterRow, MemberMapData));
    }

    #endregion
}