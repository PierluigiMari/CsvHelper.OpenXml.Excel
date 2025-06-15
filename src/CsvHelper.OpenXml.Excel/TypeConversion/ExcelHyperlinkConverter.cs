namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

/// <summary>
/// Converts an Excel hyperlink data to and from string type representations or strongly typed objects which can be <see cref="Uri"/> type or <see cref="ValueTuple{T1, T2}"/> struct where T1 and T2 are respectively of type string, string or <see cref="ValueTuple{T1, T2}"/> struct where T1 and T2 are respectively of type string, Uri.
/// </summary>
/// <remarks>This class supports various hyperlink types, such as web URLs, email addresses, and internal links, and allows customization of the resultant value type. The conversion behavior is determined by the specified <see cref="HyperlinkType"/> and <see cref="HyperlinkResultantValueType"/>.</remarks>
public class ExcelHyperlinkConverter(ExcelHyperlinkTypes hyperlinktype = ExcelHyperlinkTypes.Default, ExcelHyperlinkResultantValueTypes hyperlinkresultantvaluetype = ExcelHyperlinkResultantValueTypes.SingleString) : DefaultTypeConverter
{
    private readonly ExcelHyperlinkTypes HyperlinkType = hyperlinktype;

    private readonly ExcelHyperlinkResultantValueTypes HyperlinkResultantValueType = hyperlinktype switch
    {
        ExcelHyperlinkTypes.Default => hyperlinkresultantvaluetype == ExcelHyperlinkResultantValueTypes.SingleString ? hyperlinkresultantvaluetype : throw new ArgumentOutOfRangeException(nameof(hyperlinkresultantvaluetype), hyperlinkresultantvaluetype, "Invalid HyperlinkResultantValueType value."),
        ExcelHyperlinkTypes.WebUrl => hyperlinkresultantvaluetype,
        ExcelHyperlinkTypes.Email or ExcelHyperlinkTypes.InternalLink => hyperlinkresultantvaluetype is ExcelHyperlinkResultantValueTypes.SingleString or ExcelHyperlinkResultantValueTypes.TupleStringString ? hyperlinkresultantvaluetype : throw new ArgumentOutOfRangeException(nameof(hyperlinkresultantvaluetype), hyperlinkresultantvaluetype, "Invalid HyperlinkResultantValueType value."),
        ExcelHyperlinkTypes.FilePath => throw new NotImplementedException("FilePath HyperlinkType is not supported."),
        _ => throw new ArgumentOutOfRangeException(nameof(hyperlinktype), hyperlinktype, "Invalid ExcelHyperlinkType value.")
    };

    /// <summary>
    /// Converts a string representation of a hyperlink into an object based on the current values of HyperlinkType and HyperlinkResultantValueType.
    /// </summary>
    /// <remarks>The method processes the input string based on the <c>HyperlinkType</c> <see cref="ExcelHyperlinkTypes"/> and <c>HyperlinkResultantValueType</c>  <see cref="ExcelHyperlinkResultantValueTypes"/> properties.</remarks>
    /// <param name="text">The string representation of the hyperlink to convert.</param>
    /// <param name="row">The <see cref="IReaderRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>An object representing the converted hyperlink if the conversion succeeded; otherwise, null, or null if the value is null. The type of the returned object depends on the values of <c>HyperlinkType</c> and <c>HyperlinkResultantValueType</c>: <list type="bullet">
    /// <item><description>A <see cref="string"/> if the resultant value type is <c>SingleString</c>.</description></item>
    /// <item><description>A <see cref="Uri"/> if the resultant value type is <c>SingleUri</c>.</description></item>
    /// <item><description>A <see cref="ValueTuple{T1, T2}"/> of <see cref="string"/>, <see cref="string"/> or <see cref="string"/>, <see cref="Uri"/> respectively for TupleStringString and TupleStringUri resultant value types.</description></item> </list></returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if the combination of <c>HyperlinkType</c> and <c>HyperlinkResultantValueType</c> is invalid or unsupported.</exception>
    /// <exception cref="NotImplementedException">Thrown if the <c>HyperlinkType</c> is <c>FilePath</c>, as this type is not currently supported.</exception>
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        const string Delimiter = "(|->)";

        int DelimiterIndex = text.IndexOf(Delimiter, StringComparison.Ordinal);

        string DisplayText;
        string AddressText;

        if (DelimiterIndex >= 0)
        {
            DisplayText = text[..DelimiterIndex].Trim();
            AddressText = text[(DelimiterIndex + Delimiter.Length)..].Trim();
        }
        else
        {
            DisplayText = text.Trim();
            AddressText = text.Trim();
        }

        return HyperlinkType switch
        {
            ExcelHyperlinkTypes.Default => HyperlinkResultantValueType switch
            {
                ExcelHyperlinkResultantValueTypes.SingleString => DisplayText,
                _ => throw new ArgumentOutOfRangeException(nameof(HyperlinkResultantValueType), HyperlinkResultantValueType, "Invalid HyperlinkResultantValueType value for Default HyperlinkType.")
            },
            ExcelHyperlinkTypes.WebUrl => HyperlinkResultantValueType switch
            {
                ExcelHyperlinkResultantValueTypes.SingleString => DisplayText,
                ExcelHyperlinkResultantValueTypes.SingleUri => new Uri(AddressText),
                ExcelHyperlinkResultantValueTypes.TupleStringString => (DisplayText, AddressText),
                ExcelHyperlinkResultantValueTypes.TupleStringUri => (DisplayText, new Uri(AddressText)),
                _ => throw new ArgumentOutOfRangeException(nameof(HyperlinkResultantValueType), HyperlinkResultantValueType, "Invalid HyperlinkResultantValueType value for WebUrl HyperlinkType.")
            },
            ExcelHyperlinkTypes.Email or ExcelHyperlinkTypes.InternalLink => HyperlinkResultantValueType switch
            {
                ExcelHyperlinkResultantValueTypes.SingleString => DisplayText,
                ExcelHyperlinkResultantValueTypes.TupleStringString => (DisplayText, AddressText),
                _ => throw new ArgumentOutOfRangeException(nameof(HyperlinkResultantValueType), HyperlinkResultantValueType, "Invalid HyperlinkResultantValueType value for Email/InternalLink HyperlinkType.")
            },
            ExcelHyperlinkTypes.FilePath => throw new NotImplementedException("FilePath HyperlinkType is not supported."),
            _ => throw new ArgumentOutOfRangeException(nameof(HyperlinkType), HyperlinkType, "Invalid ExcelHyperlinkType value.")
        };
    }

    /// <summary>
    /// Converts the specified object value to its string representation based on the current values of HyperlinkType and HyperlinkResultantValueType.
    /// </summary>
    /// <remarks>The conversion behavior depends on the <c>HyperlinkType</c> <see cref="ExcelHyperlinkTypes"/> and <c>HyperlinkResultantValueType</c>  <see cref="ExcelHyperlinkResultantValueTypes"/> properties.</remarks>
    /// <param name="value">The value object to convert. This can be a string, a URI, or a ValueTuple containing a string and a URI or a string and another string.</param>
    /// <param name="row">The <see cref="IWriterRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being written.</param>
    /// <returns>A string representation of the value based on the configured <c>HyperlinkType</c> and <c>HyperlinkResultantValueType</c> if the conversion was successful; otherwise, null if the <paramref name="value"/> is null.</returns>
    /// <exception cref="NotImplementedException">Thrown if the combination of <c>HyperlinkType</c> and <c>HyperlinkResultantValueType</c> is not implemented.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if an invalid value is provided for<c>HyperlinkType</c> or <c>HyperlinkResultantValueType</c>.</exception>
    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData)
    {
        if (value is null)
        {
            return null;
        }

        const string Delimiter = "(|->)";

        return HyperlinkType switch
        {
            ExcelHyperlinkTypes.Default => value switch
            {
                ValueTuple<string, string> ss => ss.Item1,
                ValueTuple<string, Uri> su => su.Item1,
                Uri uri => uri.AbsoluteUri,
                string str => str,
                _ => throw new NotImplementedException()
            },
            ExcelHyperlinkTypes.WebUrl => HyperlinkResultantValueType switch
            {
                ExcelHyperlinkResultantValueTypes.SingleString => value switch
                {
                    ValueTuple<string, string> ss => ss.Item1,
                    ValueTuple<string, Uri> su => su.Item1,
                    Uri uri => uri.AbsoluteUri,
                    string str => str,
                    _ => throw new NotImplementedException()
                },
                ExcelHyperlinkResultantValueTypes.SingleUri => value switch
                {
                    ValueTuple<string, string> ss => ss.Item2,
                    ValueTuple<string, Uri> su => su.Item2.AbsoluteUri,
                    Uri uri => uri.AbsoluteUri,
                    string str => str,
                    _ => throw new NotImplementedException()
                },
                ExcelHyperlinkResultantValueTypes.TupleStringString => value switch
                {
                    ValueTuple<string, string> ss => string.Concat(ss.Item1, Delimiter, ss.Item2),
                    ValueTuple<string, Uri> su => string.Concat(su.Item1, Delimiter, su.Item2.AbsoluteUri),
                    Uri uri => string.Concat(uri.AbsoluteUri, Delimiter, uri.AbsoluteUri),
                    string str => string.Concat(str, Delimiter, str),
                    _ => throw new NotImplementedException()
                },
                ExcelHyperlinkResultantValueTypes.TupleStringUri => value switch
                {
                    ValueTuple<string, string> ss => string.Concat(ss.Item1, Delimiter, ss.Item2),
                    ValueTuple<string, Uri> su => string.Concat(su.Item1, Delimiter, su.Item2.AbsoluteUri),
                    Uri uri => string.Concat(uri.AbsoluteUri, Delimiter, uri.AbsoluteUri),
                    string str => string.Concat(str, Delimiter, str),
                    _ => throw new NotImplementedException()
                },
                _ => throw new ArgumentOutOfRangeException(nameof(HyperlinkResultantValueType), HyperlinkResultantValueType, "Invalid HyperlinkResultantValueType value for WebUrl HyperlinkType.")
            },
            ExcelHyperlinkTypes.Email or ExcelHyperlinkTypes.InternalLink => HyperlinkResultantValueType switch
            {
                ExcelHyperlinkResultantValueTypes.SingleString => value switch
                {
                    ValueTuple<string, string> ss => ss.Item1,
                    string str => str,
                    _ => throw new NotImplementedException()
                },
                ExcelHyperlinkResultantValueTypes.TupleStringString => value switch
                {
                    ValueTuple<string, string> ss => string.Concat(ss.Item1, Delimiter, ss.Item2),
                    string str => string.Concat(str, Delimiter, str),
                    _ => throw new NotImplementedException()
                },
                _ => throw new ArgumentOutOfRangeException(nameof(HyperlinkResultantValueType), HyperlinkResultantValueType, "Invalid HyperlinkResultantValueType value for Email/InternalLink HyperlinkType.")
            },
            ExcelHyperlinkTypes.FilePath => throw new NotImplementedException("FilePath HyperlinkType is not supported."),
            _ => throw new ArgumentOutOfRangeException(nameof(HyperlinkType), HyperlinkType, "Invalid ExcelHyperlinkType value.")
        };
    }
}