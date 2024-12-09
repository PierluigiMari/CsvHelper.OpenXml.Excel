namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

public class ExcelTimeOnlyConverter : DefaultTypeConverter
{
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData) => text is null ? null : TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(text.Replace('.', ','))));

    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : new DateOnly(1899, 12, 30).ToDateTime((TimeOnly)value).ToOADate().ToString().Replace(',', '.');
}