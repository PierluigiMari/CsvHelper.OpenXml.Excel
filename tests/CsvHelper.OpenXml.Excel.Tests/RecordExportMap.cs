namespace CsvHelper.OpenXml.Excel.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using System.Globalization;

internal class RecordExportMap : ClassMap<Record>
{
    public RecordExportMap()
    {
        AutoMap(new CultureInfo("en-US"));

        Map(m => m.Date).TypeConverter<ExcelDateTimeConverter>();
        Map(m => m.AnotherDate).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateWithDash };
        Map(m => m.YetAnotherDate).TypeConverter<ExcelDateTimeConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateWithDash };
        Map(m => m.CreationDate).TypeConverter<ExcelDateOnlyConverter>();
        Map(m => m.CreationTime).TypeConverter<ExcelTimeOnlyConverter>();
        Map(m => m.LastModifiedDate).TypeConverter<ExcelDateTimeConverter>();
    }
}