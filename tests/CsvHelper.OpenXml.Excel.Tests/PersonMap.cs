﻿namespace CsvHelper.OpenXml.Excel.Tests;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.TypeConversion;
using System.Globalization;

internal class PersonMap : ClassMap<Person>
{
    public PersonMap()
    {
        AutoMap(new CultureInfo("en-US"));

        Map(m => m.Surname).Name("Last Name");
        Map(m => m.BirthDate).Name("BirthDate").TypeConverter<ExcelDateOnlyConverter>();
        Map(m => m.Zip).Name("Zip").Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.SpecialZipCode };
        Map(m => m.CreationDate).Name("CreationDate").TypeConverter<ExcelDateOnlyConverter>();
        Map(m => m.CreationTime).TypeConverter<ExcelTimeOnlyConverter>();
        Map(m => m.LastModifiedDate).TypeConverter<ExcelDateTimeConverter>();
    }
}