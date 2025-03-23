namespace CsvHelper.OpenXml.Excel.Tests.ConsoleApp;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using System.Globalization;

public class GenericPersonMapImport : ClassMap<GenericPerson>
{
    public GenericPersonMapImport()
    {
        AutoMap(CultureInfo.CurrentCulture);
        Map(x => x.ZipCode).Convert(args => args.Row.GetField("ZipCode") is null ? string.Empty : args.Row.GetField("ZipCode")!.PadLeft(5, '0'));
        Map(x => x.DateDefault).TypeConverter<ExcelDateOnlyConverter>();
        Map(x => x.DateExtendedDefault).TypeConverter<ExcelDateOnlyConverter>();
        Map(x => x.DateWithDash).TypeConverter<ExcelDateOnlyConverter>();
        Map(x => x.GenericCurrencyWithoutDecimals);
        Map(x => x.EuroCurrencyWithTwoDecimals);
        Map(x => x.EuroCurrencyWithFourDecimals);
        Map(x => x.TimeDefault).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.TimeHoursMinutes).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.Time12HoursMinutesSeconds).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.Time12HoursMinutes).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.DateTimeDefault).TypeConverter<ExcelDateTimeConverter>();
        Map(x => x.DateTimeWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>();
        Map(x => x.DateTime12HourWithHoursMinutesSeconds).TypeConverter<ExcelDateTimeConverter>();
        Map(x => x.DateTime12HourWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>();
        //Map(m => m.DateTimeOffsetDefault).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Utc));
        //Map(m => m.DateTimeOffsetDefault).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Local));
        //Map(m => m.DateTimeOffsetDefault).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified));
        Map(m => m.DateTimeOffsetDefault);
    }
}