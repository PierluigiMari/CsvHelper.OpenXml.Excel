namespace CsvHelper.OpenXml.Excel.TypeConversion.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using Shouldly;
using Xunit;

public class ExcelDateOnlyConverterTests
{
    #region Fields

    private readonly ExcelDateOnlyConverter ExcelDateOnlyConverter = new ExcelDateOnlyConverter();
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
    [InlineData("45237", 2023, 11, 7)]
    [InlineData("45239.0", 2023, 11, 9)]
    [InlineData("45239,0", 2023, 11, 9)]
    public void ConvertFromString_ShouldConvertOADateStringToDateOnlyCorrectly(string input, int expectedyear, int expectedmonth, int expectedday)
    {
        object? ConversionResult = ExcelDateOnlyConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeOfType<DateOnly>();
        ((DateOnly)ConversionResult!).ShouldBe(new DateOnly(expectedyear, expectedmonth, expectedday));
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
    public void ConvertToString_ShouldConvertsDateOnlyToOADateStringCorrectly()
    {
        DateOnly Date = new DateOnly(2023, 11, 7);
        string ExpectedOADate = Date.ToDateTime(new TimeOnly()).ToOADate().ToString().Replace(',', '.');

        string? ConversionResult = ExcelDateOnlyConverter.ConvertToString(Date, NullWriterRow, MemberMapData);
        ConversionResult.ShouldBe(ExpectedOADate);
    }

    [Fact]
    public void ConvertToString_ShouldThrowsInvalidCastExceptionForNonDateOnly()
    {
        Should.Throw<InvalidCastException>(() => ExcelDateOnlyConverter.ConvertToString("notadate", NullWriterRow, MemberMapData));
    }

    #endregion
}