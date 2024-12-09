namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

public class ExcelDateOnlyConverter : DefaultTypeConverter
{
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData) => text is null ? null : DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(text.Replace('.', ','))));

    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : ((DateOnly)value).ToDateTime(new TimeOnly()).ToOADate().ToString().Replace(',', '.');
}