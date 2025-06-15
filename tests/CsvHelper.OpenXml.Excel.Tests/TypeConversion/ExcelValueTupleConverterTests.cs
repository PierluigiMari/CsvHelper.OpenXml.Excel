namespace CsvHelper.OpenXml.Excel.TypeConversion.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using FakeItEasy;
using Shouldly;
using Xunit;

public class ExcelValueTupleConverterTests
{
    #region Fields

    private readonly ExcelValueTupleConverter ExcelValueTupleConverter = new ExcelValueTupleConverter();
    private readonly IReaderRow NullReaderRow = null!;
    private readonly IWriterRow NullWriterRow = null!;

    #endregion

    #region Test Methods

    [Theory]
    [InlineData(null, typeof((int, string)))]
    [InlineData("", typeof((string, string)))]
    [InlineData("   ", typeof((int, string)))]
    public void ConvertFromString_ShouldReturnsNullWhenIsNullOrWhiteSpace(string? input, Type membertype)
    {
        MemberMapData MemberMapDataFake = A.Fake<MemberMapData>(x => x.WithArgumentsForConstructor(() => new MemberMapData(membertype)));
        A.CallTo(() => MemberMapDataFake.Type).Returns(membertype);

        object? ConversionResult = ExcelValueTupleConverter.ConvertFromString(input, NullReaderRow, MemberMapDataFake);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [MemberData(nameof(ConvertFromStringParameters))]
    public void ConvertFromString_ShouldConvertStringToValueTupleCorrectly(string input, Type membertype, object expectedvalue)
    {
        MemberMapData MemberMapDataFake = A.Fake<MemberMapData>(x => x.WithArgumentsForConstructor(() => new MemberMapData(membertype)));
        A.CallTo(() => MemberMapDataFake.Type).Returns(membertype);

        object? ConversionResult = ExcelValueTupleConverter.ConvertFromString(input, NullReaderRow, MemberMapDataFake);

        ConversionResult.ShouldBeOfType(membertype);
        ConversionResult.ShouldBe(expectedvalue);
    }

    [Fact]
    public void ConvertToString_ShouldReturnNullWhenIsNull()
    {
        MemberMapData MemberMapDataFake = A.Fake<MemberMapData>(x => x.WithArgumentsForConstructor(() => new MemberMapData(typeof((int,string)))));

        string? ConversionResult = ExcelValueTupleConverter.ConvertToString(null, NullWriterRow, MemberMapDataFake);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [MemberData(nameof(ConvertToStringParameters))]
    public void ConvertToString_ShouldConvertValueTupleToStringCorrectly(object? input, Type membertype, string expectedvalue)
    {
        MemberMapData MemberMapDataFake = A.Fake<MemberMapData>(x => x.WithArgumentsForConstructor(() => new MemberMapData(membertype)));
        A.CallTo(() => MemberMapDataFake.Type).Returns(membertype);

        string? ConversionResult = ExcelValueTupleConverter.ConvertToString(input, NullWriterRow, MemberMapDataFake);
        ConversionResult.ShouldBe(expectedvalue);
    }

    #endregion

    #region Parameters MemberData

    public static IEnumerable<object[]> ConvertFromStringParameters()
    {
        yield return new object[] { "42(|->)a text string", typeof((int, string)), (42, "a text string") };
        yield return new object[] { "a text string(|->)another text string", typeof((string, string)), ("a text string", "another text string") };
        yield return new object[] { "GitHub - CsvHelper.OpenXml.Excel(|->)https://github.com/PierluigiMari/CsvHelper.OpenXml.Excel", typeof((string, Uri)), ("GitHub - CsvHelper.OpenXml.Excel", new Uri("https://github.com/PierluigiMari/CsvHelper.OpenXml.Excel")) };
        yield return new object[] { "(NuGet - CsvHelper.OpenXml.Excel, https://www.nuget.org/packages/CsvHelper.OpenXml.Excel)", typeof((string, Uri)), ("NuGet - CsvHelper.OpenXml.Excel", new Uri("https://www.nuget.org/packages/CsvHelper.OpenXml.Excel")) };
        yield return new object[] { "F0756F90-B2CE-46E8-A3E1-28FBE5409A30(|->)a text string", typeof((Guid, string)), (new Guid("F0756F90-B2CE-46E8-A3E1-28FBE5409A30"), "a text string") };
        yield return new object[] { "(WebUrl, a text string)", typeof((ExcelHyperlinkTypes, string)), (ExcelHyperlinkTypes.WebUrl, "a text string") };
        yield return new object[] { "42(|->)a text string(|->)F0756F90-B2CE-46E8-A3E1-28FBE5409A30", typeof((int, string, Guid)), (42, "a text string", new Guid("F0756F90-B2CE-46E8-A3E1-28FBE5409A30")) };
    }

    public static IEnumerable<object[]> ConvertToStringParameters()
    {
        yield return new object[] { (42, "a text string"), typeof((int, string)), "(42, a text string)" };
        yield return new object[] { ("a text string", "another text string"), typeof((string, string)), "(a text string, another text string)" };
        yield return new object[] { ("GitHub - CsvHelper.OpenXml.Excel", new Uri("https://github.com/PierluigiMari/CsvHelper.OpenXml.Excel")), typeof((string, Uri)), "(GitHub - CsvHelper.OpenXml.Excel, https://github.com/PierluigiMari/CsvHelper.OpenXml.Excel)" };
        yield return new object[] { (42, "a text string", new Guid("F0756F90-B2CE-46E8-A3E1-28FBE5409A30")), typeof((int, string, Guid)), "(42, a text string, f0756f90-b2ce-46e8-a3e1-28fbe5409a30)" };
    }

    #endregion
}