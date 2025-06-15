namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;
using System.Linq;
using System.Reflection;

/// <summary>
/// Converts a string representation of a value tuple to and from <see cref="ValueTuple"/> object.
/// </summary>
public class ExcelValueTupleConverter : DefaultTypeConverter
{
    /// <summary>
    /// Converts the specified string representation of a value tuple into an <see cref="ValueTuple"/> object.
    /// </summary>
    /// <param name="text">The string to convert.</param>
    /// <param name="row">The <see cref="IReaderRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being mapped.</param>
    /// <returns>An object representing the converted ValueTuple if the conversion succeeded; or null if the value is null.</returns>
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        string[] TextValueComponents;

        if (text.Contains("(|->)", StringComparison.Ordinal))
        {
            TextValueComponents = text.Split("(|->)", StringSplitOptions.TrimEntries);
        }
        else
        {
            TextValueComponents = text[1..^1].Split(",", StringSplitOptions.TrimEntries);
        }

        Type ValueTupleType = memberMapData.Type;
        Type[] GenericsArguments = ValueTupleType.GetGenericArguments();
        int GenericsArgumentsCount = GenericsArguments.Length;

        MethodInfo CreateMethodInfo = typeof(ValueTuple).GetMethods().First(m => m.Name == "Create" && m.GetGenericArguments().Length == GenericsArgumentsCount);

        object[] ConvertedValueComponents = new object[GenericsArgumentsCount];

        for (int i = 0; i < GenericsArgumentsCount; i++)
        {
            string TextValueComponent = TextValueComponents[i];
            Type TargetType = GenericsArguments[i];

            if (TargetType == typeof(Uri))
                ConvertedValueComponents[i] = new Uri(TextValueComponent, UriKind.RelativeOrAbsolute);
            else if (TargetType == typeof(Guid))
                ConvertedValueComponents[i] = Guid.Parse(TextValueComponent);
            else if (TargetType.IsEnum)
                ConvertedValueComponents[i] = Enum.Parse(TargetType, TextValueComponent);
            else
                ConvertedValueComponents[i] = Convert.ChangeType(TextValueComponent, TargetType);
        }

        return CreateMethodInfo.MakeGenericMethod(GenericsArguments).Invoke(null, ConvertedValueComponents);
    }

    /// <summary>
    /// Converts the specified <see cref="ValueTuple"/> object to its string representation.
    /// </summary>
    /// <param name="value">The <see cref="ValueTuple"/> object to convert.</param>
    /// <param name="row">The <see cref="IWriterRow"/> for the current record.</param>
    /// <param name="memberMapData">The <see cref="MemberMapData"/> for the member being written.</param>
    /// <returns>A string representation of the  <see cref="ValueTuple"/> object if the conversion was successful; otherwise, null if the <paramref name="value"/> is null.</returns>
    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : base.ConvertToString(value, row, memberMapData);
}