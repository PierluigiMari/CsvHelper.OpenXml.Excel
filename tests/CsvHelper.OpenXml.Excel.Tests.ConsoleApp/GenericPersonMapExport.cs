namespace CsvHelper.OpenXml.Excel.Tests.ConsoleApp;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using System.Globalization;

public class GenericPersonMapExport : ClassMap<GenericPerson>
{
    public GenericPersonMapExport()
    {
        AutoMap(CultureInfo.CurrentCulture);
        Map(x => x.NumberInteger).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.NumberIntegerDefault };
        Map(x => x.NumberDecimalWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.NumberDecimalWithTwoDecimalsDefault };
        Map(x => x.NumberDecimalWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.NumberDecimalWithFourDecimals };

        Map(x => x.PhoneNo).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Text };

        //Map(x => x.EmailId).TypeConverter(new ExcelHyperlinkConverter());//.Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };

        Map(x => x.ZipCode).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.SpecialZipCode };

        Map(x => x.DateDefault).TypeConverter<ExcelDateOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateDefault };
        Map(x => x.DateExtendedDefault).TypeConverter<ExcelDateOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateExtended };
        Map(x => x.DateWithDash).TypeConverter<ExcelDateOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateWithDash };

        Map(x => x.GenericCurrencyWithoutDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyGenericWithoutDecimals };
        Map(x => x.GenericCurrencyWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyGenericWithTwoDecimals };
        Map(x => x.GenericCurrencyWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyGenericWithFourDecimals };

        Map(x => x.EuroCurrencyWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyEuroITWithTwoDecimals };
        Map(x => x.EuroCurrencyWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyEuroITWithFourDecimals };

        Map(x => x.AccountingWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.AccountingEuroITWithTwoDecimals };
        Map(x => x.AccountingWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.AccountingEuroITWithFourDecimals };

        Map(x => x.Percentage).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.PercentageWithoutDecimals };
        Map(x => x.PercentageWithTwoDecimal).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.PercentageWithTwoDecimals };

        Map(x => x.TimeDefault).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.TimeWithHoursMinutesSecondsDefault };
        Map(x => x.TimeHoursMinutes).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.TimeWithHoursMinutes };
        Map(x => x.Time12HoursMinutesSeconds).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Time12HourWithHoursMinutesSeconds };
        Map(x => x.Time12HoursMinutes).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Time12HourWithHoursMinutes };

        Map(x => x.DateTimeDefault).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutes };
        Map(x => x.DateTime12HourWithHoursMinutesSeconds).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTime12HourWithHoursMinutesSeconds };
        Map(x => x.DateTime12HourWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTime12HourWithHoursMinutes };

        Map(m => m.DateTimeOffsetUnspecified).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetUnspecifiedAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        Map(x => x.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        //Map(x => x.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = new CultureInfo("en-US") };
        //Map(m => m.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(m => m.DateTimeOffsetUtc).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Utc)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetUtcAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Utc)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        Map(m => m.DateTimeOffsetLocal).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Local)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetLocalAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Local)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };

        Map(x => x.ScientificDefault).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.ScientificWithTwoDecimalsDefault };
        Map(x => x.ScientificFourDecimal).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.ScientificWithFourDecimals };

        //Map(x => x.Link).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        //Map(x => x.LinkText).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
    }

    public GenericPersonMapExport(CsvContext csvcontext)
    {
        AutoMap(csvcontext);
        Map(x => x.NumberInteger).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.NumberIntegerDefault };
        Map(x => x.NumberDecimalWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.NumberDecimalWithTwoDecimalsDefault };
        Map(x => x.NumberDecimalWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.NumberDecimalWithFourDecimals };

        Map(x => x.PhoneNo).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Text };

        //Map(x => x.EmailId).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.TupleStringString)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };

        Map(x => x.ZipCode).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.SpecialZipCode };

        Map(x => x.DateDefault).TypeConverter<ExcelDateOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateDefault };
        Map(x => x.DateExtendedDefault).TypeConverter<ExcelDateOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateExtended };
        Map(x => x.DateWithDash).TypeConverter<ExcelDateOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateWithDash };

        Map(x => x.GenericCurrencyWithoutDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyGenericWithoutDecimals };
        Map(x => x.GenericCurrencyWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyGenericWithTwoDecimals };
        Map(x => x.GenericCurrencyWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyGenericWithFourDecimals };

        Map(x => x.EuroCurrencyWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyEuroITWithTwoDecimals };
        Map(x => x.EuroCurrencyWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.CurrencyEuroITWithFourDecimals };

        Map(x => x.AccountingWithTwoDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.AccountingEuroITWithTwoDecimals };
        Map(x => x.AccountingWithFourDecimals).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.AccountingEuroITWithFourDecimals };

        Map(x => x.Percentage).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.PercentageWithoutDecimals };
        Map(x => x.PercentageWithTwoDecimal).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.PercentageWithTwoDecimals };

        Map(x => x.TimeDefault).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.TimeWithHoursMinutesSecondsDefault };
        Map(x => x.TimeHoursMinutes).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.TimeWithHoursMinutes };
        Map(x => x.Time12HoursMinutesSeconds).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Time12HourWithHoursMinutesSeconds };
        Map(x => x.Time12HoursMinutes).TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Time12HourWithHoursMinutes };

        Map(x => x.DateTimeDefault).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutes };
        Map(x => x.DateTime12HourWithHoursMinutesSeconds).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTime12HourWithHoursMinutesSeconds };
        Map(x => x.DateTime12HourWithHoursMinutes).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTime12HourWithHoursMinutes };

        Map(m => m.DateTimeOffsetUnspecified).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetUnspecifiedAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        Map(x => x.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        //Map(x => x.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = new CultureInfo("en-US") };
        //Map(m => m.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(m => m.DateTimeOffsetUtc).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Utc)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetUtcAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Utc)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        Map(m => m.DateTimeOffsetLocal).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Local)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetLocalAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Local)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };

        Map(x => x.ScientificDefault).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.ScientificWithTwoDecimalsDefault };
        Map(x => x.ScientificFourDecimal).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.ScientificWithFourDecimals };

        //Map(x => x.LinkText).TypeConverter(new ExcelValueTupleConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        //Map(x => x.LinkText).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        //Map(x => x.Link).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        //Map(x => x.LinkToCell).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.InternalLink, ExcelHyperlinkResultantValueTypes.TupleStringString)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        //Map(x => x.LinkToCell).TypeConverter(new ExcelValueTupleConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
    }
}