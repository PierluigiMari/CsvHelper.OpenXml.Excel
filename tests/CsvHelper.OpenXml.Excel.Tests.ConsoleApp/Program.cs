// See https://aka.ms/new-console-template for more information
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel;
using CsvHelper.OpenXml.Excel.Tests.ConsoleApp;
using System.Globalization;
using System.Text.Json;

Console.WriteLine("- Verifications Console App -");
Console.WriteLine();

foreach (VerificationEnumeration verification in Enum.GetValues<VerificationEnumeration>())
{
    Console.WriteLine($"{(int)verification} - {verification.ToString()}");
}

int? VerificationToPerformValue;

do
{
    Console.WriteLine();

    Console.WriteLine("What verification do you intend to perform?");

    string? VerificationToPerformInput = Console.ReadLine();

    VerificationToPerformValue = VerificationToPerformInput is null ? 0 : int.TryParse(VerificationToPerformInput, out int VerificationToPerformInt) ? VerificationToPerformInt : null;

    if (VerificationToPerformValue is null)
    {
        Console.WriteLine();

        Console.WriteLine("Invalid verification entered. Please try again.");
    }
} while (VerificationToPerformValue is null);

VerificationEnumeration VerificationToPerform = (VerificationEnumeration)VerificationToPerformValue;

bool Result = VerificationToPerform switch
{
    VerificationEnumeration.DomImportExcelFile => DomImportExcelFile(),
    VerificationEnumeration.DomImportExcelFileAsyncEnumerable => await DomImportExcelFileAsyncEnumerable(),
    VerificationEnumeration.DomImportExcelFileDynamic => DomImportExcelFileDynamic(),
    VerificationEnumeration.DomImportExcelFileDynamicAsyncEnumerable => await DomImportExcelFileDynamicAsyncEnumerable(),
    VerificationEnumeration.SaxImportExcelFile => SaxImportExcelFile(),
    VerificationEnumeration.SaxImportExcelFileAsyncEnumerable => await SaxImportExcelFileAsyncEnumerable(),
    VerificationEnumeration.SaxImportExcelFileDynamic => SaxImportExcelFileDynamic(),
    VerificationEnumeration.SaxImportExcelFileDynamicAsyncEnumerable => await SaxImportExcelFileDynamicAsyncEnumerable(),
    VerificationEnumeration.DomExportExcelFile => DomExportExcelFile(),
    VerificationEnumeration.DomExportExcelFileTwoCollectionSameTypeOnSameSheet => DomExportExcelFileTwoCollectionSameTypeOnSameSheet(),
    VerificationEnumeration.DomExportExcelFileTwoCollectionSameTypeOnDifferentSheet => DomExportExcelFileTwoCollectionSameTypeOnDifferentSheet(),
    VerificationEnumeration.DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheet => DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheet(),
    VerificationEnumeration.DomExportExcelFileAsync => await DomExportExcelFileAsync(),
    VerificationEnumeration.DomExportExcelFileTwoCollectionSameTypeOnSameSheetAsync => await DomExportExcelFileTwoCollectionSameTypeOnSameSheetAsync(),
    VerificationEnumeration.DomExportExcelFileTwoCollectionSameTypeOnDifferentSheetAsync => await DomExportExcelFileTwoCollectionSameTypeOnDifferentSheetAsync(),
    VerificationEnumeration.DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheetAsync => await DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheetAsync(),

    _ => false
};


bool DomImportExcelFile()
{
    byte[] ExcelBytes = File.ReadAllBytes("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");

    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    ExcelReader.Context.RegisterClassMap<GenericPersonMapImport>();

    IEnumerable<GenericPerson> People = ExcelReader.GetRecords<GenericPerson>().ToArray();

    Console.WriteLine();

    foreach (GenericPerson person in People)
    {
        Console.WriteLine(person.ToString());

        Console.WriteLine();
    }

    return true;
}

async Task<bool> DomImportExcelFileAsyncEnumerable()
{
    byte[] ExcelBytes = await File.ReadAllBytesAsync("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");

    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    ExcelReader.Context.RegisterClassMap<GenericPersonMapImport>();

    IAsyncEnumerable<GenericPerson> PeopleAsync = ExcelReader.GetRecordsAsync<GenericPerson>();

    Console.WriteLine();

    await foreach (GenericPerson person in PeopleAsync)
    {
        Console.WriteLine(person.ToString());

        Console.WriteLine();
    }

    return true;
}

bool DomImportExcelFileDynamic()
{
    byte[] ExcelBytes = File.ReadAllBytes("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");
    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    IEnumerable<dynamic> People = ExcelReader.GetRecords<dynamic>().ToArray();

    Console.WriteLine();

    foreach (dynamic person in People)
    {
        Console.WriteLine(JsonSerializer.Serialize(person));
        Console.WriteLine();
    }

    return true;
}

async Task<bool> DomImportExcelFileDynamicAsyncEnumerable()
{
    byte[] ExcelBytes = await File.ReadAllBytesAsync("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");
    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    IAsyncEnumerable<dynamic> PeopleAsync = ExcelReader.GetRecordsAsync<dynamic>();

    Console.WriteLine();

    await foreach (dynamic person in PeopleAsync)
    {
        Console.WriteLine(JsonSerializer.Serialize(person));
        Console.WriteLine();
    }

    return true;
}

bool SaxImportExcelFile()
{
    byte[] ExcelBytes = File.ReadAllBytes("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");

    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    ExcelReader.Context.RegisterClassMap<GenericPersonMapImport>();

    IEnumerable<GenericPerson> People = ExcelReader.GetRecords<GenericPerson>().ToArray();

    Console.WriteLine();

    foreach (GenericPerson person in People)
    {
        Console.WriteLine(person.ToString());

        Console.WriteLine();
    }

    return true;
}

async Task<bool> SaxImportExcelFileAsyncEnumerable()
{
    byte[] ExcelBytes = await File.ReadAllBytesAsync("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");

    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    ExcelReader.Context.RegisterClassMap<GenericPersonMapImport>();

    IAsyncEnumerable<GenericPerson> PeopleAsync = ExcelReader.GetRecordsAsync<GenericPerson>();

    Console.WriteLine();

    await foreach (GenericPerson person in PeopleAsync)
    {
        Console.WriteLine(person.ToString());

        Console.WriteLine();
    }

    return true;
}

bool SaxImportExcelFileDynamic()
{
    byte[] ExcelBytes = File.ReadAllBytes("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");
    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    IEnumerable<dynamic> People = ExcelReader.GetRecords<dynamic>().ToArray();

    Console.WriteLine();

    foreach (dynamic person in People)
    {
        Console.WriteLine(JsonSerializer.Serialize(person));
        Console.WriteLine();
    }

    return true;
}

async Task<bool> SaxImportExcelFileDynamicAsyncEnumerable()
{
    byte[] ExcelBytes = await File.ReadAllBytesAsync("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx");
    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Foglio1");
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    IAsyncEnumerable<dynamic> PeopleAsync = ExcelReader.GetRecordsAsync<dynamic>();

    Console.WriteLine();

    await foreach (dynamic person in PeopleAsync)
    {
        Console.WriteLine(JsonSerializer.Serialize(person));
        Console.WriteLine();
    }

    return true;
}

bool DomExportExcelFile()
{
    IEnumerable<GenericPerson> People = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        ExcelWriter.WriteRecords<GenericPerson>(People, "People");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    File.WriteAllBytes($"{OutputDirectoryInfo.FullName}/DomExportExcelFile.xlsx", ExcelFile);

    return true;
}

bool DomExportExcelFileTwoCollectionSameTypeOnSameSheet()
{
    IEnumerable<GenericPerson> People1 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    IEnumerable<GenericPerson> People2 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio2");

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        ExcelWriter.WriteRecords<GenericPerson>(People1, "People");

        ExcelWriter.NextRecord();

        ExcelWriter.WriteRecords<GenericPerson>(People2, "People");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    File.WriteAllBytes($"{OutputDirectoryInfo.FullName}/DomExportExcelFileTwoCollectionSameTypeOnSameSheet.xlsx", ExcelFile);

    return true;
}

bool DomExportExcelFileTwoCollectionSameTypeOnDifferentSheet()
{
    IEnumerable<GenericPerson> People1 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    IEnumerable<GenericPerson> People2 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio2");

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        ExcelWriter.WriteRecords<GenericPerson>(People1, "People1");

        ExcelWriter.Context.UnregisterClassMap<GenericPersonMapExport>();

        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        ExcelWriter.WriteRecords<GenericPerson>(People2, "People2");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    File.WriteAllBytes($"{OutputDirectoryInfo.FullName}/DomExportExcelFileTwoCollectionSameTypeOnDifferentSheet.xlsx", ExcelFile);

    return true;
}

bool DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheet()
{
    IEnumerable<GenericPerson> People = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    IEnumerable<GenericOrder> Orders = new List<GenericOrder>
    {
        new GenericOrder
        {
            OrderId = 1,
            OrderNumber = "ORD-2021001",
            OrderDate = new DateOnly(2021, 1, 2),
            OrderTime = new TimeOnly(12, 0),
            OrderAmount = 100.50m,
            CustomerName = "John Doe",
            CustomerAddress = "250 Via Tuscolana",
            CustomerZip = "00181",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2021, 1, 3, 10, 25, 15)
        },
        new GenericOrder
        {
            OrderId = 5,
            OrderNumber = "ORD-2021005",
            OrderDate = new DateOnly(2021, 1, 3),
            OrderTime = new TimeOnly(14, 5),
            OrderAmount = 150.80m,
            CustomerName = "Alexander Green",
            CustomerAddress = "100 Via Merulana",
            CustomerZip = "00185",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2021, 1, 4, 10, 15, 05)
        },
        new GenericOrder
        {
            OrderId = 50,
            OrderNumber = "ORD-2021050",
            OrderDate = new DateOnly(2021, 1, 31),
            OrderTime = new TimeOnly(10, 45),
            OrderAmount = 320.30m,
            CustomerName = "Ethan Clarke",
            CustomerAddress = "100 Via Palmiro Togliatti",
            CustomerZip = "00155",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2021, 2, 4, 14, 10, 45)
        },
        new GenericOrder
        {
            OrderId = 500,
            OrderNumber = "ORD-2023500",
            OrderDate = new DateOnly(2023, 5, 25),
            OrderTime = new TimeOnly(10, 0),
            OrderAmount = 200.75m,
            CustomerName = "Jane Doe",
            CustomerAddress = "250 Via Tuscolana",
            CustomerZip = "00181",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2023, 5, 26, 9, 25, 15)
        }
    };

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        ExcelWriter.WriteRecords<GenericPerson>(People, "People");

        ExcelWriter.Context.UnregisterClassMap<GenericPersonMapExport>();

        ExcelWriter.Context.RegisterClassMap<GenericOrderMapExport>();

        ExcelWriter.WriteRecords<GenericOrder>(Orders, "Orders");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    File.WriteAllBytes($"{OutputDirectoryInfo.FullName}/DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheet.xlsx", ExcelFile);

    return true;
}

async Task<bool> DomExportExcelFileAsync()
{
    IEnumerable<GenericPerson> People = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        await ExcelWriter.WriteRecordsAsync<GenericPerson>(People, "People");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    await File.WriteAllBytesAsync($"{OutputDirectoryInfo.FullName}/DomExportExcelFileAsync.xlsx", ExcelFile);

    return true;
}

async Task<bool> DomExportExcelFileTwoCollectionSameTypeOnSameSheetAsync()
{
    IEnumerable<GenericPerson> People1 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    IEnumerable<GenericPerson> People2 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio2");

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        await ExcelWriter.WriteRecordsAsync<GenericPerson>(People1, "People");

        ExcelWriter.NextRecord();

        await ExcelWriter.WriteRecordsAsync<GenericPerson>(People2, "People");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    await File.WriteAllBytesAsync($"{OutputDirectoryInfo.FullName}/DomExportExcelFileTwoCollectionSameTypeOnSameSheetAsync.xlsx", ExcelFile);

    return true;
}

async Task<bool> DomExportExcelFileTwoCollectionSameTypeOnDifferentSheetAsync()
{
    IEnumerable<GenericPerson> People1 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    IEnumerable<GenericPerson> People2 = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio2");

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        await ExcelWriter.WriteRecordsAsync<GenericPerson>(People1, "People1");

        ExcelWriter.Context.UnregisterClassMap<GenericPersonMapExport>();

        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        await ExcelWriter.WriteRecordsAsync<GenericPerson>(People2, "People2");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    await File.WriteAllBytesAsync($"{OutputDirectoryInfo.FullName}/DomExportExcelFileTwoCollectionSameTypeOnDifferentSheetAsync.xlsx", ExcelFile);

    return true;
}

async Task<bool> DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheetAsync()
{
    IEnumerable<GenericPerson> People = ImportGenericPersonExcelFileForExport("InputTestFiles/CsvHelperOpenXmlInputTest.xlsx", "Foglio1");

    IEnumerable<GenericOrder> Orders = new List<GenericOrder>
    {
        new GenericOrder
        {
            OrderId = 1,
            OrderNumber = "ORD-2021001",
            OrderDate = new DateOnly(2021, 1, 2),
            OrderTime = new TimeOnly(12, 0),
            OrderAmount = 100.50m,
            CustomerName = "John Doe",
            CustomerAddress = "250 Via Tuscolana",
            CustomerZip = "00181",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2021, 1, 3, 10, 25, 15)
        },
        new GenericOrder
        {
            OrderId = 5,
            OrderNumber = "ORD-2021005",
            OrderDate = new DateOnly(2021, 1, 3),
            OrderTime = new TimeOnly(14, 5),
            OrderAmount = 150.80m,
            CustomerName = "Alexander Green",
            CustomerAddress = "100 Via Merulana",
            CustomerZip = "00185",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2021, 1, 4, 10, 15, 05)
        },
        new GenericOrder
        {
            OrderId = 50,
            OrderNumber = "ORD-2021050",
            OrderDate = new DateOnly(2021, 1, 31),
            OrderTime = new TimeOnly(10, 45),
            OrderAmount = 320.30m,
            CustomerName = "Ethan Clarke",
            CustomerAddress = "100 Via Palmiro Togliatti",
            CustomerZip = "00155",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2021, 2, 4, 14, 10, 45)
        },
        new GenericOrder
        {
            OrderId = 500,
            OrderNumber = "ORD-2023500",
            OrderDate = new DateOnly(2023, 5, 25),
            OrderTime = new TimeOnly(10, 0),
            OrderAmount = 200.75m,
            CustomerName = "Jane Doe",
            CustomerAddress = "250 Via Tuscolana",
            CustomerZip = "00181",
            CustomerCity = "Roma",
            ShippedDate = new DateTime(2023, 5, 26, 9, 25, 15)
        }
    };

    using MemoryStream ExcelStream = new MemoryStream();
    using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CsvConfiguration(CultureInfo.CurrentCulture)))
    {
        ExcelWriter.Context.RegisterClassMap<GenericPersonMapExport>();

        await ExcelWriter.WriteRecordsAsync<GenericPerson>(People, "People");

        ExcelWriter.Context.UnregisterClassMap<GenericPersonMapExport>();

        ExcelWriter.Context.RegisterClassMap<GenericOrderMapExport>();

        await ExcelWriter.WriteRecordsAsync<GenericOrder>(Orders, "Orders");
    }

    byte[] ExcelFile = ExcelStream.ToArray();

    DirectoryInfo OutputDirectoryInfo = Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/OutputTestFiles");

    await File.WriteAllBytesAsync($"{OutputDirectoryInfo.FullName}/DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheetAsync.xlsx", ExcelFile);

    return true;
}


IEnumerable<GenericPerson> ImportGenericPersonExcelFileForExport(string fileName, string sheetname)
{
    byte[] ExcelBytes = File.ReadAllBytes(fileName);

    using MemoryStream ExcelStream = new MemoryStream(ExcelBytes);
    using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, sheetname);
    using CsvReader ExcelReader = new CsvReader(ExcelParser);

    ExcelReader.Context.RegisterClassMap<GenericPersonMapImport>();

    return ExcelReader.GetRecords<GenericPerson>().ToArray();
}