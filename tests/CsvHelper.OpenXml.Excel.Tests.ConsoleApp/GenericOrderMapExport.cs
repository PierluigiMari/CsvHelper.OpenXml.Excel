namespace CsvHelper.OpenXml.Excel.Tests.ConsoleApp;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using System.Globalization;


public class GenericOrderMapExport : ClassMap<GenericOrder>
{
    public GenericOrderMapExport()
    {
        AutoMap(CultureInfo.CurrentCulture);
        Map(m => m.OrderDate).TypeConverter<ExcelDateOnlyConverter>();
        Map(m => m.OrderTime).TypeConverter<ExcelTimeOnlyConverter>();
        Map(m => m.CustomerZip).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.SpecialZipCode };
        Map(m => m.ShippedDate).TypeConverter<ExcelDateTimeConverter>();
    }
}