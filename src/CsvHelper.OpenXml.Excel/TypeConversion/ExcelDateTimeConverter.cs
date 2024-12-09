namespace CsvHelper.OpenXml.Excel.TypeConversion;

using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;

public class ExcelDateTimeConverter : DefaultTypeConverter
{
    public override object? ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData) => text is null ? null : DateTime.FromOADate(double.Parse(text.Replace('.', ',')));

    public override string? ConvertToString(object? value, IWriterRow row, MemberMapData memberMapData) => value is null ? null : ((DateTime)value).ToOADate().ToString().Replace(',', '.');
}