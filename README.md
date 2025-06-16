# CsvHelper.OpenXml.Excel

[![NuGet](https://img.shields.io/nuget/v/CsvHelper.OpenXml.Excel.svg)](https://www.nuget.org/packages/CsvHelper.OpenXml.Excel)

## Introduction
**CsvHelper.OpenXml.Excel** is a small library thinked to allow the import and export data from and to **Excel** files. The library is the facto an extension that connects and integrates two famous libraries [CsvHelper](https://github.com/JoshClose/CsvHelper) and [OpenXml](https://github.com/dotnet/Open-XML-SDK). It mainly providing implementations of `IParser` and `IWriter`, of **CsvHelper**, which read and write files in xlsx format using **OpenXml**.

The ultimate goal is to obtain versatility of use and accuracy in import and export results; especially with regard to export, the file obtained, although always in simple tabular form, still has all the characteristics expected for an Excel file, with the columns having the cells formatted in an adequate way and not as simple text.

## Prerequisites
Knowledge of [CsvHelper](https://github.com/JoshClose/CsvHelper) and its documentation.

.NET 8 or .NET 9 SDK installed.

## Installation
To install the library, from the **Package Manager Console** use the following command:

```
PM> Install-Package CsvHelper.OpenXml.Excel
```
Or,  from the **.NET Core CLI Console** use the following command:

```
 > dotnet add package CsvHelper.OpenXml.Excel
```

## Import
For importing from Excel files, the library makes available both approaches offered by **OpenXml**, so there are two implementations of **IParser**:

- ExcelDomParser
- ExcelSaxParser

Both perform the exact same task, with the same input parameters and configuration characteristics, of course the second is strongly recommended for importing very large files, to avoid Out of Memory exceptions.

### Example usage
**ExcelDomParser** and **ExcelSaxParser** can be used by specifying an instance of Stream as the principal constructor parameter.

```csharp
byte[] Bytes = File.ReadAllBytes("path/subpath/file.xlsx");

using MemoryStream ExcelStream = new MemoryStream(Bytes);     
using var ExcelParser = new ExcelDomParser(ExcelStream);
using var ExcelReader = new CsvReader(ExcelParser);

IEnumerable<Foo> FooCollection = ExcelReader.GetRecords<Foo>().ToArray();
```

>The constructor has two optional parameters, **sheetname**, which allows you to specify the name of the worksheet; **configuration**, for which the instance of a `CsvConfiguration` is used. In case the sheet name is not specified, the first worksheet is used as the data source by default; if the configuration is not specified, a configuration with InvariantCulture is used by default.

The library, also, provides specific implementations of `DefaultTypeConverter` that can be used in the definition of `ClassMap` and that allow, for Excel files having columns with specific cell formats (**Date** or **Time** or **Custom (Date and Time)**), to define mappings to the **DateOnly** or **TimeOnly** or **DateTime** types.

- ExcelDateOnlyConverter
- ExcelTimeOnlyConverter
- ExcelDateTimeConverter

The use of these converters is very versatile, in fact, for example, it's possible to define the mapping of an Excel column having cell format **Custom (Date and Time)** to a **DateOnly** type, but also to a **TimeOnly** type; or define the mapping of an Excel column with a **Time** cell format to a **DateTime** type; as even define the mapping of an Excel column having format **Date** to a **DateTime** type.

```csharp
public class FooMap : ClassMap<Foo>
{
    public FooMap()
    {
        AutoMap(CultureInfo.CurrentCulture);
        Map(x => x.Date).TypeConverter<ExcelDateOnlyConverter>();
        Map(x => x.Time).TypeConverter<ExcelTimeOnlyConverter>();
        Map(x => x.DateTime).TypeConverter<ExcelDateTimeConverter>();
    }
}
```

A separate discussion must be made for the following converters, which allow, for Excel files having columns with specific cell formats **(Custom (Date and Time)**, **Custom (Date and Time) in text format** or **Custom (Date and Time) in text format with the information of the offset with respect to UTC time**), to define the mappings to the **DateTimeOffset** type.

- ExcelDateTimeOffsetConverter
- ExcelDateTimeOffsetTextConverter

With the use of these converters it's possible to define the mapping of an Excel column having cell format **Custom (Date and Time)** to a **DateTimeOffset** type using an `ExcelDateTimeOffsetConverter`; or map an Excel column having cell format **Text** with the value of a **Custom (Date and Time)** or with the value of a **Custom (Date and Time) followed by the offset information with respect to UTC time**, to a **DateTimeOffset** type using an `ExcelDateTimeOffsetTextConverter`.

```csharp
public class FooMap : ClassMap<Foo>
{
    public FooMap()
    {
        AutoMap(CultureInfo.CurrentCulture);

        Map(m => m.DateTimeOffsetUnspecified).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified, new TimeSpan(3, 0, 0)));
        Map(m => m.DateTimeOffsetUnspecifiedAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter());

        Map(m => m.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Unspecified, new TimeSpan(3, 0, 0)));

        Map(m => m.DateTimeOffsetUtc).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Utc, new TimeSpan(0, 0, 0)));
        Map(m => m.DateTimeOffsetUtcAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter());

        Map(m => m.DateTimeOffsettUtcFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Utc, new TimeSpan(0, 0, 0)));

        Map(m => m.DateTimeOffsetLocal).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Local, new TimeSpan(1, 0, 0)));
        Map(m => m.DateTimeOffsetLocalAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter());
        
        Map(m => m.DateTimeOffsetLocalFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Local, new TimeSpan(1, 0, 0)));
    }
}
```

```csharp
byte[] Bytes = File.ReadAllBytes("path/subpath/file.xlsx");

using MemoryStream ExcelStream = new MemoryStream(Bytes);     
using var ExcelParser = new ExcelDomParser(ExcelStream);
using var ExcelReader = new CsvReader(ExcelParser);

ExcelReader.Context.RegisterClassMap<FooMap>();

IEnumerable<Foo> FooCollection = ExcelReader.GetRecords<Foo>().ToArray();
```

The library, also, provides specific implementations that can be used in the definition of `ClassMap` and that allow, for Excel files having columns with specific Hyperlink cell formats or columns with specific text cell formats attributable to a ValueTuple, to define the mapping to the ValueTuple type.

- ExcelValueTupleConverter
- ExcelHyperlink Converter

Because CsvHelper doesn't have a DefaultTypeConverter for the ValueTuple type, we had to define an ExcelValueTupleConverter, and then we defined the hyperlink-specific converter, ExcelHyperlinkConverter.

```csharp
public class FooMap : ClassMap<Foo>
{
    public FooMap()
    {
        AutoMap(CultureInfo.CurrentCulture);
        Map(m => m.Hyperlink).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.SingleUri));
        Map(m => m.HyperlinkWithText).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri));
        Map(m => m.HyperlinkEmail).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.TupleStringString));
        Map(m => m.HyperlinkToAnotherCell).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.InternalLink, ExcelHyperlinkResultantValueTypes.TupleStringString));
    }
}
```

>Regarding import, always keep in mind that it's the value of the cell that is imported, regardless of its formatting; in Excel you can have a cell formatted **Currency (2 decimals)**, but its value could be with 4 decimals, so, in this specific case, assuming that the destination type is **Decimal**, the imported data would be with 4 decimals. If it's necessary to obtain an imported data in the destination type that is more faithful to what is displayed on Excel, it will be necessary to intervene in the definition of the `ClassMap` by specifying in the mapping the type of conversion with rounding that is to be applied for that specific Excel column.
>
>Assuming that Foo has a Price property of type **Decimal**, and the Excel file has a Price column with cell formatting **Currency (2 decimals)** but with a value of 4 decimals, to get the above said
>```csharp
>Map(x => x.Price).Convert(args => args.Row.GetField("Price") is null ? 0 : Math.Round(decimal.Parse(args.Row.GetField("Price")!.Replace('.', ',')), 2));
>```

## Export
For exporting to Excel files, the library currently only provides the DOM approach of the two offered by OpenXml, so for now there is only one implementation of IWriter:

- ExcelDomWriter

### Example usage
**ExcelDomWriter**, like **ExcelDomParser** and **ExcelSaxParser**, can be used by specifying an instance of Stream as the principal constructor parameter.

```csharp
using MemoryStream ExcelStream = new MemoryStream();
using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
{
    ExcelWriter.WriteRecords(FooCollection);
}

byte[] Bytes = ExcelStream.ToArray();

File.WriteAllBytes("path/subpath/file.xlsx", Bytes);
```

>The constructor has an optional parameter, **configuration**, for which the instance of a `CsvConfiguration` is used. In case the  configuration is not specified, a configuration with InvariantCulture is used by default.

Like with import, you can use the specific implementations of the `DefaultTypeConverter` to the definition of `ClassMap`. The library, also, provides both an enumeration:

- ExcelCellFormats

which allows you to specify the Excel cell format to be applied to the column of the generated worksheet, and a specific implementation of `TypeConverterOptions`:

- ExcelTypeConverterOptions

that adds the Excel cell format, to the options that can be used to define the type conversion.

```csharp
public class FooMap : ClassMap<Foo>
{
    public FooMap()
    {
        AutoMap(CultureInfo.CurrentCulture);
        Map(x => x.Date).TypeConverter<ExcelDateOnlyConverter>()
            .Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateDefault };
        Map(x => x.Time).TypeConverter<ExcelTimeOnlyConverter>()
            .TypeConverter<ExcelTimeOnlyConverter>().Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.TimeHoursMinutesSecondsDefault };
        Map(x => x.DateTime).TypeConverter<ExcelDateTimeConverter>()
            .Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeDefault };

        //With version => 1.2.0
        Map(m => m.DateTimeOffsetUnspecified).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetUnspecifiedAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        
        Map(x => x.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        // Or even
        Map(m => m.DateTimeOffsetUnspecifiedFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Unspecified)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        
        Map(m => m.DateTimeOffsetUtc).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Utc)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault  };
        Map(x => x.DateTimeOffsetUtcAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Utc)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        
        Map(x => x.DateTimeOffsetUtcFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };

        Map(m => m.DateTimeOffsetLocal).TypeConverter(new ExcelDateTimeOffsetConverter(DateTimeKind.Local)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault };
        Map(x => x.DateTimeOffsetLocalAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter(DateTimeKind.Local)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
        
        Map(x => x.DateTimeOffsetLocalFromDateTimeAsText).TypeConverter(new ExcelDateTimeOffsetTextConverter()).Data.TypeConverterOptions = new ExcelTypeConverterOptions { CultureInfo = CultureInfo.CurrentCulture };
    }
}
```
```csharp
using MemoryStream ExcelStream = new MemoryStream();
using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
{
    ExcelWriter.Context.RegisterClassMap<FooMap>();
    ExcelWriter.WriteRecords(FooCollection);
}

byte[] Bytes = ExcelStream.ToArray();

File.WriteAllBytes("path/subpath/file.xlsx", Bytes);
```

**ExcelCellFormats** enumeration, the following are the details of the defined named constant members:

0. **Default** :arrow_right: Default format
0. **DefaultBold** :arrow_right: Format with Bold text
0. **DefaultBoldCentered** :arrow_right: Format with Bold text with horizontal aligment centered
0. **NumberIntegerDefault** :arrow_right: Format like "0" - ***Default*** for "**Int32**" type
0. **NumberDecimalWithTwoDecimalsDefault** :arrow_right: Format like "0.00" - ***Default*** for "**Decimal**" type
0. **NumberDecimalWithFourDecimals** :arrow_right: Format like "0.0000"
0. **CurrencyGenericWithoutDecimals** :arrow_right: Format like "#,##0"
0. **CurrencyGenericWithTwoDecimals** :arrow_right: Format like "#,##0.00"
0. **CurrencyGenericWithFourDecimals** :arrow_right: Format like "#,##0.0000"
0. **CurrencyEuroITWithTwoDecimals** :arrow_right: Format like "#,##0.00 &euro;"
0. **CurrencyEuroITWithFourDecimals** :arrow_right: Format like "#,##0.0000 &euro;"
0. **CurrencyDollarUSWithTwoDecimals** :arrow_right: Format like "&dollar;#,##0.00"
0. **CurrencyDollarUSWithFourDecimals** :arrow_right: Format like "&dollar;#,##0.0000"
0. **CurrencyPoundGBWithTwoDecimals** :arrow_right: Format like "&pound;#,##0.00"
0. **CurrencyPoundGBWithFourDecimals** :arrow_right: Format like "&pound;#,##0.0000"
0. **AccountingEuroITWithTwoDecimals** :arrow_right: Format like "#,##0.00 &euro;"
0. **AccountingEuroITWithFourDecimals** :arrow_right: Format like "#,##0.0000 &euro;"
0. **AccountingDollarUSWithTwoDecimals** :arrow_right: Format like "&dollar; #,##0.00"
0. **AccountingDollarUSWithFourDecimals** :arrow_right: Format like "&dollar; #,##0.0000"
0. **AccountingPoundGBWithTwoDecimals** :arrow_right: Format like "&pound; #,##0.00;"
0. **AccountingPoundGBWithFourDecimals** :arrow_right: Format like "&pound; #,##0.0000"
0. **DateDefault** :arrow_right: Format like "dd/mm/yyyy" - ***Default*** for "**DateOnly**" type
0. **DateExtended** :arrow_right: Format like "dddd dd mmmm yyyy"
0. **DateWithDash** :arrow_right: Format like "dd-mm-yyyy"
0. **DateTimeWithHoursMinutesSecondsDefault** :arrow_right: Format like "dd/mm/yyyy hh:mm:ss" - ***Default*** for "**DateTime**" type
0. **DateTimeWithHoursMinutes** :arrow_right: Format like "dd/mm/yyyy hh:mm"
0. **DateTime12HourWithHoursMinutesSeconds** :arrow_right: Format like "dd/mm/yyyy h:mm:ss AM/PM"
0. **DateTime12HourWithHoursMinutes** :arrow_right: Format like "dd/mm/yyyy h:mm AM/PM"
0. **TimeWithHoursMinutesSecondsDefault** :arrow_right: Format like "hh:mm:ss" - ***Default*** for "**TimeOnly**" type
0. **TimeWithHoursMinutes** :arrow_right: Format like "hh:mm"
0. **Time12HourWithHoursMinutesSeconds** :arrow_right: Format like "h:mm:ss AM/PM"
0. **Time12HourWithHoursMinutes** :arrow_right: Format like "h:mm AM/PM"
0. **PercentageWithoutDecimals** :arrow_right: Format like "0%"
0. **PercentageWithTwoDecimals** :arrow_right: Format like "0.00%"
0. **ScientificWithTwoDecimalsDefault** :arrow_right: Format like "0.00E+00" - ***Default*** for "**Double**" type
0. **ScientificWithFourDecimals** :arrow_right: Format like "0.0000E+00"
0. **Text** :arrow_right: Format like plain text
0. **SpecialZipCode** :arrow_right: Format like "00000"

>For each type (Int32, DateOnly, DateTime, TimeOnly, Double) a default enumerate has been defined, recognizable by "...Default" at the end of the name; in the `ClassMap` definition, can omit the Excel cell format if intend to apply the default format to that Excel column.
>
>Assuming that the Date property, of Foo, is of type DateOnly, the member Map can also be written in the following way
>```csharp
>Map(x => x.Date).TypeConverter<ExcelDateOnlyConverter>();
>
>//Instead of
>Map(x => x.Date).TypeConverter<ExcelDateOnlyConverter>()
>    .Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.DateDefault };
>```

For the following converters, in addition to a separate discussion, a slightly different approach must also be used with regard to the definition of the ```ClassMap``` and as regards the use, we are talking about the converters that allow, for ValueTuple type properties to define the mapping to obtain Excel files having columns with specific Hyperlink cell formats or columns with specific Text cell formats attributable to a ValueTuple.

- ExcelValueTupleConverter
- ExcelHyperlinkConverter

Since CsvHelper not have a DefaultTypeConverter for the ValueTuple type, it was necessary to define an ExcelValueTupleConverter, so the specific converter for Hyperlinks, the ExcelHyperlinkConverter, was defined.

```csharp
public class FooMap : ClassMap<Foo>
{
    public FooMap(CsvContext csvcontext)
    {
        AutoMap(csvcontext);
        Map(x => x.Hyperlink).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        Map(x => x.HyperlinkWithText).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.WebUrl, ExcelHyperlinkResultantValueTypes.TupleStringUri)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        Map(x => x.HyperlinkEmail).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.Email, ExcelHyperlinkResultantValueTypes.TupleStringString)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };
        Map(x => x.HyperlinkToAnotherCell).TypeConverter(new ExcelHyperlinkConverter(ExcelHyperlinkTypes.InternalLink, ExcelHyperlinkResultantValueTypes.TupleStringString)).Data.TypeConverterOptions = new ExcelTypeConverterOptions { ExcelCellFormat = ExcelCellFormats.Hyperlink };

        //Or if you want to obtain a string representation of the ValueTuple in the cell.
        Map(x => x.HyperlinkWithText).TypeConverter(new ExcelValueTupleConverter());
    }
}
```
```csharp
using MemoryStream ExcelStream = new MemoryStream();
using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
{
    //The following AddConverters are required to enable the recognition of the ValueTuple type by CsvHelper which otherwise would not recognize them,
    //also causing the incorrect placement of the specific columns in the Excel file.
    ExcelWriter.Context.TypeConverterCache.AddConverter<ValueTuple<string, Uri>>(new ExcelValueTupleConverter());
    ExcelWriter.Context.TypeConverterCache.AddConverter<ValueTuple<string, string>>(new ExcelValueTupleConverter());

    //This differentiation is necessary to force CsvHelper to maintain context, which otherwise would not be considered by him.
    ExcelWriter.Context.RegisterClassMap(new FooMap(ExcelWriter.Context));
    ExcelWriter.WriteRecords(FooCollection);
}

byte[] Bytes = ExcelStream.ToArray();

File.WriteAllBytes("path/subpath/file.xlsx", Bytes);
```

The library, also, allows the export of **Enumeration types**, by default for Enum type properties the named constants are exported (the names to be clear), if you want to export the value you have to proceed by specifying it in the definition of the ClassMap.

>Assuming that Foo has a Season property of type Seasons enum
>```csharp
>enum Seasons : int
>{
>    Spring = 0,
>    Summer = 1,
>    Autumn = 2,
>    Winter = 3
>}
>```
>To export the value of the enum assigned to the Season of Foo property, the Member Map can be written as follows
>```csharp
>Map(m => m.Season).TypeConverterOption.Format("D");
>```
>To be clear, for `Season = Seasons.Summer` in the Excel cell would be exported `1`
>
>To export the name of the enumeration assigned to the Season of Foo property while still specifying it in the ClassMap definition, the member map can be written as follows
>```csharp
>Map(m => m.Season).TypeConverterOption.Format("G");
>```
>To be clear, for `Season = Seasons.Summer` in the Excel cell would be exported `Summer`

### Usage details
If you have two collections of the same type, you do not need to concat them into a single collection to proceed with the export.

```csharp
using MemoryStream ExcelStream = new MemoryStream();
using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
{
    ExcelWriter.Context.RegisterClassMap<FooMap>();
    ExcelWriter.WriteRecords(FooCollection);
    
    ExcelWriter.NextRecord();

    ExcelWriter.WriteRecords(AnotherFooCollection);
}

byte[] Bytes = ExcelStream.ToArray();

File.WriteAllBytes("path/subpath/file.xlsx", Bytes);
```

>This puts, the data in the second collection in the same worksheet as the first collection.

If you have two collections of different types and you need to export them both to the same Excel file.

```csharp
using MemoryStream ExcelStream = new MemoryStream();
using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
{
    ExcelWriter.Context.RegisterClassMap<FooMap>();
    ExcelWriter.WriteRecords(FooCollection, "SheetFoo");

    ExcelWriter.Context.UnregisterClassMap<FooMap>();

    ExcelWriter.Context.RegisterClassMap<BarMap>();
    ExcelWriter.WriteRecords(BarCollection, "SheetBar");
}

byte[] Bytes = ExcelStream.ToArray();

File.WriteAllBytes("path/subpath/file.xlsx", Bytes);
```

>This puts the data in the second collection in a different worksheet than the first collection.