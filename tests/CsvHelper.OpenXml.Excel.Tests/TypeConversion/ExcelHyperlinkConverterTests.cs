namespace CsvHelper.OpenXml.Excel.TypeConversion.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using FakeItEasy;
using Shouldly;
using Xunit;

public class ExcelHyperlinkConverterTests
{
    #region Fields

    private readonly MemberMapData MemberMapData = new MemberMapData(null);
    private readonly IReaderRow NullReaderRow = null!;
    private readonly IWriterRow NullWriterRow = null!;

    #endregion

    #region Test Methods


    [Theory]
    [InlineData(null, typeof((int, string)))]
    [InlineData("", typeof((string, string)))]
    [InlineData("   ", typeof((int, string)))]
    public void ConvertFromString_ShouldReturnNullWhenIsNullOrWhitespace(string? input, Type membertype)
    {
        MemberMapData MemberMapDataFake = A.Fake<MemberMapData>(x => x.WithArgumentsForConstructor(() => new MemberMapData(membertype)));
        A.CallTo(() => MemberMapDataFake.Type).Returns(membertype);

        ExcelHyperlinkConverter ExcelHyperlinkConverter = new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleString);
        object? ConversionResult = ExcelHyperlinkConverter.ConvertFromString(input, NullReaderRow, MemberMapDataFake);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [MemberData(nameof(ConvertFromStringWebUrlParameters))]
    [MemberData(nameof(ConvertFromStringEmailParameters))]
    [MemberData(nameof(ConvertFromStringInternalLinkParameters))]
    [MemberData(nameof(ConvertFromStringDefaultParameters))]
    public void ConvertFromString_ShouldConvertStringToValueTupleCorrectly(string input, ExcelHyperlinkTypes hyperlinktype, ExcelHyperlinkResultantValueTypes resultanttype, Type expectedtype, object expectedvalue)
    {
        ExcelHyperlinkConverter ExcelHyperlinkConverter = new ExcelHyperlinkConverter(hyperlinktype, resultanttype);

        object? ConversionResult = ExcelHyperlinkConverter.ConvertFromString(input, NullReaderRow, MemberMapData);

        ConversionResult.ShouldBeOfType(expectedtype);
        ConversionResult.ShouldBe(expectedvalue);
    }

    [Theory]
    [InlineData(ExcelHyperlinkTypes.WebUrl, (ExcelHyperlinkResultantValueTypes)999)]
    public void ConvertFromString_ShouldThrowsArgumentOutOfRangeExceptionForInvalidResultantType(ExcelHyperlinkTypes hyperlinktype, ExcelHyperlinkResultantValueTypes resultanttype)
    {
        ExcelHyperlinkConverter ExcelHyperlinkConverter = new ExcelHyperlinkConverter(hyperlinktype, resultanttype);

        Should.Throw<ArgumentOutOfRangeException>(() => ExcelHyperlinkConverter.ConvertFromString("test", NullReaderRow, MemberMapData));
    }

    [Fact]
    public void ConvertToString_ShouldReturnNullWhenIsNull()
    {
        ExcelHyperlinkConverter ExcelHyperlinkConverter = new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleString);

        string? ConversionResult = ExcelHyperlinkConverter.ConvertToString(null, NullWriterRow, MemberMapData);

        ConversionResult.ShouldBeNull();
    }

    [Theory]
    [MemberData(nameof(ConvertToStringWebUrlParameters))]
    [MemberData(nameof(ConvertToStringEmailParameters))]
    [MemberData(nameof(ConvertToStringInternalLinkParameters))]
    [MemberData(nameof(ConvertToStringDefaultParameters))]
    public void ConvertToString_ShouldConvertValueTupleToStyringCorrectly(ExcelHyperlinkTypes hyperlinktype, ExcelHyperlinkResultantValueTypes resultanttype, object value, string expected)
    {
        ExcelHyperlinkConverter ExcelHyperlinkConverter = new ExcelHyperlinkConverter(hyperlinktype, resultanttype);

        string? ConversionResult = ExcelHyperlinkConverter.ConvertToString(value, NullWriterRow, MemberMapData);

        ConversionResult.ShouldBe(expected);
    }


    [Fact]
    public void ConvertToString_ShouldThrowsArgumentOutOfRangeExceptionForInvalidResultantType()
    {
        ExcelHyperlinkConverter ExcelHyperlinkConverter = new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, (ExcelHyperlinkResultantValueTypes)999);
        
        Should.Throw<ArgumentOutOfRangeException>(() =>  ExcelHyperlinkConverter.ConvertToString("test", NullWriterRow, MemberMapData));
    }

    #endregion

    #region Parameters MemberData

    public static IEnumerable<object[]> ConvertFromStringWebUrlParameters()
    {
        yield return new object[] { "Google(|->)https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleString, typeof(string), "Google" };
        yield return new object[] { "Google(|->)https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleUri, typeof(Uri), new Uri("https://www.google.com") };
        yield return new object[] { "Google(|->)https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringString, typeof((string, string)), ("Google", "https://www.google.com") };
        yield return new object[] { "Google(|->)https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri, typeof((string, Uri)), ("Google", new Uri("https://www.google.com")) };
        yield return new object[] { "https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleString, typeof(string), "https://www.google.com" };
        yield return new object[] { "https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleUri, typeof(Uri), new Uri("https://www.google.com") };
        yield return new object[] { "https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringString, typeof((string, string)), ("https://www.google.com", "https://www.google.com") };
        yield return new object[] { "https://www.google.com", ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri, typeof((string, Uri)), ("https://www.google.com", new Uri("https://www.google.com")) };
    }

    public static IEnumerable<object[]> ConvertFromStringEmailParameters()
    {
        yield return new object[] { "Contact(|->)mailto:someone@example.com", ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.SingleString, typeof(string), "Contact" };
        yield return new object[] { "Contact(|->)mailto:someone@example.com", ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.TupleStringString, typeof((string, string)), ("Contact", "mailto:someone@example.com") };
        yield return new object[] { "mailto:someone@example.com", ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.SingleString, typeof(string), "mailto:someone@example.com" };
        yield return new object[] { "mailto:someone@example.com", ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.TupleStringString, typeof((string, string)), ("mailto:someone@example.com", "mailto:someone@example.com") };
    }
    
    public static IEnumerable<object[]> ConvertFromStringInternalLinkParameters()
    {
        yield return new object[] { "Sheet1!A1(|->)#Sheet1!A1", ExcelHyperlinkTypes.InternalLink, ExcelHyperlinkResultantValueTypes.SingleString, typeof(string), "Sheet1!A1" };
        yield return new object[] { "Sheet1!A1(|->)#Sheet1!A1", ExcelHyperlinkTypes.InternalLink, ExcelHyperlinkResultantValueTypes.TupleStringString, typeof((string, string)), ("Sheet1!A1", "#Sheet1!A1") };
    }

    public static IEnumerable<object[]> ConvertFromStringDefaultParameters()
    {
        yield return new object[] { "Just a string", ExcelHyperlinkTypes.Default, ExcelHyperlinkResultantValueTypes.SingleString, typeof(string), "Just a string" };
        yield return new object[] { "  Just a string  ", ExcelHyperlinkTypes.Default, ExcelHyperlinkResultantValueTypes.SingleString, typeof(string), "Just a string" };
    }

    public static IEnumerable<object[]> ConvertToStringWebUrlParameters()
    {
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleString, ("Google", "https://www.google.com"), "Google" };
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleUri, ("Google", new Uri("https://www.google.com")), "https://www.google.com/" };
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringString, ("Google", "https://www.google.com"), "Google(|->)https://www.google.com" };
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri, ("Google", new Uri("https://www.google.com")), "Google(|->)https://www.google.com/" };
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleString, "Google", "Google" };
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleUri, new Uri("https://www.google.com"), "https://www.google.com/" };
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringString, "Google", "Google(|->)Google" };
        yield return new object[] { ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri, "Google", "Google(|->)Google" };
    }

    public static IEnumerable<object[]> ConvertToStringEmailParameters()
    {
        yield return new object[] { ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.SingleString, ("Contact", "mailto:someone@example.com"), "Contact" };
        yield return new object[] { ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.TupleStringString, ("Contact", "mailto:someone@example.com"), "Contact(|->)mailto:someone@example.com" };
        yield return new object[] { ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.SingleString, "mailto:someone@example.com", "mailto:someone@example.com" };
        yield return new object[] { ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.TupleStringString, "mailto:someone@example.com", "mailto:someone@example.com(|->)mailto:someone@example.com" };
    }

    public static IEnumerable<object[]> ConvertToStringInternalLinkParameters()
    {
        yield return new object[] { ExcelHyperlinkTypes.InternalLink, ExcelHyperlinkResultantValueTypes.SingleString, ("Sheet1!A1", "#Sheet1!A1"), "Sheet1!A1" };
        yield return new object[] { ExcelHyperlinkTypes.InternalLink, ExcelHyperlinkResultantValueTypes.TupleStringString, ("Sheet1!A1", "#Sheet1!A1"), "Sheet1!A1(|->)#Sheet1!A1" };
    }

    public static IEnumerable<object[]> ConvertToStringDefaultParameters()
    {
        yield return new object[] { ExcelHyperlinkTypes.Default, ExcelHyperlinkResultantValueTypes.SingleString, ("Just a string", "ignored"), "Just a string" };
        yield return new object[] { ExcelHyperlinkTypes.Default, ExcelHyperlinkResultantValueTypes.SingleString, "Just a string", "Just a string" };
        yield return new object[] { ExcelHyperlinkTypes.Default, ExcelHyperlinkResultantValueTypes.SingleString, new Uri("https://www.google.com"), "https://www.google.com/" };
    }

    #endregion
}