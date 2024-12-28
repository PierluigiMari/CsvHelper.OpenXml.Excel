namespace CsvHelper.OpenXml.Excel.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using System.Globalization;

internal class OrderMap : ClassMap<Order>
{
    public OrderMap()
    {
        AutoMap(new CultureInfo("en-US"));
        Map(m => m.OrderDate).TypeConverter<ExcelDateOnlyConverter>();
        Map(m => m.OrderTime).TypeConverter<ExcelTimeOnlyConverter>();
        Map(m => m.CustomerZip).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.SpecialZipCode };
        Map(m => m.ShippedDate).TypeConverter<ExcelDateTimeConverter>();
    }
}