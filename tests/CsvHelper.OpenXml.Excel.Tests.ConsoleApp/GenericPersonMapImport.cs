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
        Map(x => x.TimeDefault).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.TimeHoursMinutes).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.Time12HoursMinutesSeconds).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.Time12HoursMinutes).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.DateTimeDefault).TypeConverter<ExcelDateTimeConverter>();
        Map(x => x.DateTimeWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>();
        Map(x => x.DateTime12HourWithHoursMinutesSeconds).TypeConverter<ExcelDateTimeConverter>();
        Map(x => x.DateTime12HourWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>();

        Map(m => m.DateTimeOffsetUnspecified).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified, new TimeSpan(3, 0, 0)));
        Map(m => m.DateTimeOffsetUnspecifiedAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter());

        Map(m => m.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Unspecified, new TimeSpan(4, 0, 0)));

        Map(m => m.DateTimeOffsetUtc).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Utc, new TimeSpan(0, 0, 0)));
        Map(m => m.DateTimeOffsetUtcAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter());

        Map(m => m.DateTimeOffsetLocal).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Local, new TimeSpan(1, 0, 0)));
        Map(m => m.DateTimeOffsetLocalAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter());
    }
}