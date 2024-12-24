namespace CsvHelper.OpenXml.Excel.Tests;

using CsvHelper.OpenXml.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using System;
using System.Dynamic;
using System.Globalization;
using System.Threading.Tasks;
using Xunit;

public class ExcelDomWriterTests
{
    #region Test Methods

    [Fact()]
    public void WriteSingleAnonymousUnformattedRecordTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            var Record = new
            {
                Name = "John",
                Surname = "Doe",
                NickName = null as string,
                BirthDate = new DateOnly(1994, 1, 6),
                Age = 30,
                Address = "250 Via Tuscolana",
                Zip = "00181",
                City = "Roma",
                CreationDate = new DateOnly(2021, 1, 2),
                CreationTime = new TimeOnly(12, 0),
                LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
            };

            ExcelWriter.WriteRecord(Record);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(2);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");

    }

    [Fact()]
    public void WriteMultipleAnonymousUnformattedRecordsTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            var Records = new[]
            {
                new
                {
                    Name = "John",
                    Surname = "Doe",
                    NickName = null as string,
                    BirthDate = new DateOnly(1994, 1, 6),
                    Age = 30,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2021, 1, 2),
                    CreationTime = new TimeOnly(12, 0),
                    LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
                },
                new
                {
                    Name = "Jane",
                    Surname = "Doe",
                    NickName = (string?)"Tarzan lady",
                    BirthDate = new DateOnly(1996, 3, 15),
                    Age = 28,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2023, 5, 25),
                    CreationTime = new TimeOnly(10, 0),
                    LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
                }
            }.ToList();

            ExcelWriter.WriteRecords(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public async Task WriteMultipleAnonymousUnformattedRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            var Records = new[]
            {
                new
                {
                    Name = "John",
                    Surname = "Doe",
                    NickName = null as string,
                    BirthDate = new DateOnly(1994, 1, 6),
                    Age = 30,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2021, 1, 2),
                    CreationTime = new TimeOnly(12, 0),
                    LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
                },
                new
                {
                    Name = "Jane",
                    Surname = "Doe",
                    NickName = (string?)"Tarzan lady",
                    BirthDate = new DateOnly(1996, 3, 15),
                    Age = 28,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2023, 5, 25),
                    CreationTime = new TimeOnly(10, 0),
                    LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
                }
            }.ToList();

            await ExcelWriter.WriteRecordsAsync(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public async Task WriteMultipleAnonymousUnformattedAsyncRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            await ExcelWriter.WriteRecordsAsync(GetAnonymousAsyncEnumerable());
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public void WriteSingleDynamicUnformattedRecordTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            dynamic Record = new ExpandoObject();
            Record.Name = "John";
            Record.Surname = "Doe";
            Record.NickName = null as string;
            Record.BirthDate = new DateOnly(1994, 1, 6);
            Record.Age = 30;
            Record.Address = "250 Via Tuscolana";
            Record.Zip = "00181";
            Record.City = "Roma";
            Record.CreationDate = new DateOnly(2021, 1, 2);
            Record.CreationTime = new TimeOnly(12, 0);
            Record.LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15));

            ExcelWriter.WriteRecord(Record);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(2);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public void WriteMultipleDynamicUnformattedRecordsTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            dynamic Record1 = new ExpandoObject();
            Record1.Name = "John";
            Record1.Surname = "Doe";
            Record1.NickName = null as string;
            Record1.BirthDate = new DateOnly(1994, 1, 6);
            Record1.Age = 30;
            Record1.Address = "250 Via Tuscolana";
            Record1.Zip = "00181";
            Record1.City = "Roma";
            Record1.CreationDate = new DateOnly(2021, 1, 2);
            Record1.CreationTime = new TimeOnly(12, 0);
            Record1.LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15));

            dynamic Record2 = new ExpandoObject();
            Record2.Name = "Jane";
            Record2.Surname = "Doe";
            Record2.NickName = (string?)"Tarzan lady";
            Record2.BirthDate = new DateOnly(1996, 3, 15);
            Record2.Age = 28;
            Record2.Address = "250 Via Tuscolana";
            Record2.Zip = "00181";
            Record2.City = "Roma";
            Record2.CreationDate = new DateOnly(2023, 5, 25);
            Record2.CreationTime = new TimeOnly(10, 0);
            Record2.LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15));

            List<dynamic> Records = new List<dynamic> { Record1, Record2 };

            ExcelWriter.WriteRecords(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public async Task WriteMultipleDynamicUnformattedRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            dynamic Record1 = new ExpandoObject();
            Record1.Name = "John";
            Record1.Surname = "Doe";
            Record1.NickName = null as string;
            Record1.BirthDate = new DateOnly(1994, 1, 6);
            Record1.Age = 30;
            Record1.Address = "250 Via Tuscolana";
            Record1.Zip = "00181";
            Record1.City = "Roma";
            Record1.CreationDate = new DateOnly(2021, 1, 2);
            Record1.CreationTime = new TimeOnly(12, 0);
            Record1.LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15));

            dynamic Record2 = new ExpandoObject();
            Record2.Name = "Jane";
            Record2.Surname = "Doe";
            Record2.NickName = (string?)"Tarzan lady";
            Record2.BirthDate = new DateOnly(1996, 3, 15);
            Record2.Age = 28;
            Record2.Address = "250 Via Tuscolana";
            Record2.Zip = "00181";
            Record2.City = "Roma";
            Record2.CreationDate = new DateOnly(2023, 5, 25);
            Record2.CreationTime = new TimeOnly(10, 0);
            Record2.LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15));

            List<dynamic> Records = new List<dynamic> { Record1, Record2 };

            await ExcelWriter.WriteRecordsAsync(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public async Task WriteMultipleDynamicUnformattedAsyncRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            await ExcelWriter.WriteRecordsAsync(GetDynamicAsyncEnumerable());
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public void WriteSinglePersonUnformattedRecordTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            Person Record = new Person
            {
                Name = "John",
                Surname = "Doe",
                NickName = null as string,
                BirthDate = new DateOnly(1994, 1, 6),
                Age = 30,
                Address = "250 Via Tuscolana",
                Zip = "00181",
                City = "Roma",
                CreationDate = new DateOnly(2021, 1, 2),
                CreationTime = new TimeOnly(12, 0),
                LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
            };

            ExcelWriter.WriteRecord(Record);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(2);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");

    }

    [Fact()]
    public void WriteMultiplePersonUnformattedRecordsTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            List<Person> Records = new List<Person>
            {
                new Person
                {
                    Name = "John",
                    Surname = "Doe",
                    NickName = null as string,
                    BirthDate = new DateOnly(1994, 1, 6),
                    Age = 30,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2021, 1, 2),
                    CreationTime = new TimeOnly(12, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                },
                new Person
                {
                    Name = "Jane",
                    Surname = "Doe",
                    NickName = (string?)"Tarzan lady",
                    BirthDate = new DateOnly(1996, 3, 15),
                    Age = 28,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2023, 5, 25),
                    CreationTime = new TimeOnly(10, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                }
            };

            ExcelWriter.WriteRecords(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public async Task WriteMultiplePersonUnformattedRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            List<Person> Records = new List<Person>
            {
                new Person
                {
                    Name = "John",
                    Surname = "Doe",
                    NickName = null as string,
                    BirthDate = new DateOnly(1994, 1, 6),
                    Age = 30,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2021, 1, 2),
                    CreationTime = new TimeOnly(12, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                },
                new Person
                {
                    Name = "Jane",
                    Surname = "Doe",
                    NickName = (string?)"Tarzan lady",
                    BirthDate = new DateOnly(1996, 3, 15),
                    Age = 28,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2023, 5, 25),
                    CreationTime = new TimeOnly(10, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                }
            };

            await ExcelWriter.WriteRecordsAsync(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public async Task WriteMultiplePersonUnformattedAsyncRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            await ExcelWriter.WriteRecordsAsync(GetPersonAsyncEnumerable());
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Surname");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().Be("1/6/1994");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Should().Be("3/15/1996");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7)).Should().Be("1/2/2021");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8)).Should().Be("5/25/2023");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Should().Be("12:00 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Should().Be("10:00 AM");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Should().Be("12/24/2024 3:25:15 PM");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Should().Be("12/24/2024 3:25:15 PM");
    }

    [Fact()]
    public void WriteSinglePersonFormattedRecordTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            ExcelWriter.Context.RegisterClassMap<PersonMap>();

            Person Record = new Person
            {
                Name = "John",
                Surname = "Doe",
                NickName = null as string,
                BirthDate = new DateOnly(1994, 1, 6),
                Age = 30,
                Address = "250 Via Tuscolana",
                Zip = "00181",
                City = "Roma",
                CreationDate = new DateOnly(2021, 1, 2),
                CreationTime = new TimeOnly(12, 0),
                LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
            };

            ExcelWriter.WriteRecord(Record);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(2);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Last Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Replace('.', ',')))).Should().Be(new DateOnly(1994, 1, 6));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7))))).Should().Be(new DateOnly(2021, 1, 2));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Replace('.', ',')))).Should().Be(new TimeOnly(12, 0, 0));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Replace('.', ','))).Should().Be(new DateTime(2024, 12, 24, 15, 25, 15));
    }

    [Fact()]
    public void WriteMultiplePersonFormattedRecordsTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            ExcelWriter.Context.RegisterClassMap<PersonMap>();

            List<Person> Records = new List<Person>
            {
                new Person
                {
                    Name = "John",
                    Surname = "Doe",
                    NickName = null as string,
                    BirthDate = new DateOnly(1994, 1, 6),
                    Age = 30,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2021, 1, 2),
                    CreationTime = new TimeOnly(12, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                },
                new Person
                {
                    Name = "Jane",
                    Surname = "Doe",
                    NickName = (string?)"Tarzan lady",
                    BirthDate = new DateOnly(1996, 3, 15),
                    Age = 28,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2023, 5, 25),
                    CreationTime = new TimeOnly(10, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                }
            };

            ExcelWriter.WriteRecords(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Last Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Replace('.', ',')))).Should().Be(new DateOnly(1994, 1, 6));
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Replace('.', ',')))).Should().Be(new DateOnly(1996, 3, 15));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7))))).Should().Be(new DateOnly(2021, 1, 2));
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8))))).Should().Be(new DateOnly(2023, 5, 25));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Replace('.', ',')))).Should().Be(new TimeOnly(12, 0, 0));
        TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Replace('.', ',')))).Should().Be(new TimeOnly(10, 0, 0));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Replace('.', ','))).Should().Be(new DateTime(2024, 12, 24, 15, 25, 15));
        DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Replace('.', ','))).Should().Be(new DateTime(2024, 12, 24, 15, 25, 15));
    }

    [Fact()]
    public async Task WriteMultiplePersonFormattedRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            ExcelWriter.Context.RegisterClassMap<PersonMap>();

            List<Person> Records = new List<Person>
            {
                new Person
                {
                    Name = "John",
                    Surname = "Doe",
                    NickName = null as string,
                    BirthDate = new DateOnly(1994, 1, 6),
                    Age = 30,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2021, 1, 2),
                    CreationTime = new TimeOnly(12, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                },
                new Person
                {
                    Name = "Jane",
                    Surname = "Doe",
                    NickName = (string?)"Tarzan lady",
                    BirthDate = new DateOnly(1996, 3, 15),
                    Age = 28,
                    Address = "250 Via Tuscolana",
                    Zip = "00181",
                    City = "Roma",
                    CreationDate = new DateOnly(2023, 5, 25),
                    CreationTime = new TimeOnly(10, 0),
                    LastModifiedDate = new DateTime(2024, 12, 24, 15, 25, 15)
                }
            };

            await ExcelWriter.WriteRecordsAsync(Records);
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Last Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Replace('.', ',')))).Should().Be(new DateOnly(1994, 1, 6));
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Replace('.', ',')))).Should().Be(new DateOnly(1996, 3, 15));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7))))).Should().Be(new DateOnly(2021, 1, 2));
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8))))).Should().Be(new DateOnly(2023, 5, 25));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Replace('.', ',')))).Should().Be(new TimeOnly(12, 0, 0));
        TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Replace('.', ',')))).Should().Be(new TimeOnly(10, 0, 0));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Replace('.', ','))).Should().Be(new DateTime(2024, 12, 24, 15, 25, 15));
        DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Replace('.', ','))).Should().Be(new DateTime(2024, 12, 24, 15, 25, 15));
    }

    [Fact()]
    public async Task WriteMultiplePersonFormattedAsyncRecordsAsyncTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using (ExcelDomWriter ExcelWriter = new ExcelDomWriter(ExcelStream, new CultureInfo("en-US")))
        {
            ExcelWriter.Context.RegisterClassMap<PersonMap>();

            await ExcelWriter.WriteRecordsAsync(GetPersonAsyncEnumerable());
        }

        ExcelStream.Position = 0;
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Open(ExcelStream, false);
        SheetData? ExcelSheetData = ExcelDocument.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
        List<Row> ExcelRows = ExcelSheetData!.Elements<Row>().ToList();

        ExcelRows.Should().HaveCount(3);
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(0)).Should().Be("Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(0)).Should().Be("John");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(0)).Should().Be("Jane");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(1)).Should().Be("Last Name");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(1)).Should().Be("Doe");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(2)).Should().Be("NickName");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Should().NotBe("");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(2)).Should().Be("Tarzan lady");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(3)).Should().Be("BirthDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(2)).Replace('.', ',')))).Should().Be(new DateOnly(1994, 1, 6));
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(3)).Replace('.', ',')))).Should().Be(new DateOnly(1996, 3, 15));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(4)).Should().Be("Age");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(3)).Should().Be("30");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(4)).Should().Be("28");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(5)).Should().Be("Address");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(4)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(5)).Should().Be("250 Via Tuscolana");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(6)).Should().Be("Zip");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(5)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(6)).Should().Be("00181");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(7)).Should().Be("City");
        GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(6)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(7)).Should().Be("Roma");
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(8)).Should().Be("CreationDate");
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(7))))).Should().Be(new DateOnly(2021, 1, 2));
        DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(8))))).Should().Be(new DateOnly(2023, 5, 25));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(9)).Should().Be("CreationTime");
        TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(8)).Replace('.', ',')))).Should().Be(new TimeOnly(12, 0, 0));
        TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(9)).Replace('.', ',')))).Should().Be(new TimeOnly(10, 0, 0));
        GetCellValue(ExcelDocument, ExcelRows[0].Elements<Cell>().ElementAt(10)).Should().Be("LastModifiedDate");
        DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[1].Elements<Cell>().ElementAt(9)).Replace('.', ','))).Should().Be(new DateTime(2024, 12, 24, 15, 25, 15));
        DateTime.FromOADate(double.Parse(GetCellValue(ExcelDocument, ExcelRows[2].Elements<Cell>().ElementAt(10)).Replace('.', ','))).Should().Be(new DateTime(2024, 12, 24, 15, 25, 15));
    }

    #endregion

    #region Private Methods

    private async IAsyncEnumerable<Person> GetPersonAsyncEnumerable()
    {
        yield return new Person
        {
            Name = "John",
            Surname = "Doe",
            NickName = null as string,
            BirthDate = new DateOnly(1994, 1, 6),
            Age = 30,
            Address = "250 Via Tuscolana",
            Zip = "00181",
            City = "Roma",
            CreationDate = new DateOnly(2021, 1, 2),
            CreationTime = new TimeOnly(12, 0),
            LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
        };
        ;
        yield return new Person
        {
            Name = "Jane",
            Surname = "Doe",
            NickName = (string?)"Tarzan lady",
            BirthDate = new DateOnly(1996, 3, 15),
            Age = 28,
            Address = "250 Via Tuscolana",
            Zip = "00181",
            City = "Roma",
            CreationDate = new DateOnly(2023, 5, 25),
            CreationTime = new TimeOnly(10, 0),
            LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
        };

        await Task.CompletedTask;
    }

    private async IAsyncEnumerable<object> GetAnonymousAsyncEnumerable()
    {
        yield return new
        {
            Name = "John",
            Surname = "Doe",
            NickName = null as string,
            BirthDate = new DateOnly(1994, 1, 6),
            Age = 30,
            Address = "250 Via Tuscolana",
            Zip = "00181",
            City = "Roma",
            CreationDate = new DateOnly(2021, 1, 2),
            CreationTime = new TimeOnly(12, 0),
            LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
        };
;
        yield return new
        {
            Name = "Jane",
            Surname = "Doe",
            NickName = (string?)"Tarzan lady",
            BirthDate = new DateOnly(1996, 3, 15),
            Age = 28,
            Address = "250 Via Tuscolana",
            Zip = "00181",
            City = "Roma",
            CreationDate = new DateOnly(2023, 5, 25),
            CreationTime = new TimeOnly(10, 0),
            LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15))
        };

        await Task.CompletedTask;
    }

    private async IAsyncEnumerable<dynamic> GetDynamicAsyncEnumerable()
    {
        dynamic Record1 = new ExpandoObject();
        Record1.Name = "John";
        Record1.Surname = "Doe";
        Record1.NickName = null as string;
        Record1.BirthDate = new DateOnly(1994, 1, 6);
        Record1.Age = 30;
        Record1.Address = "250 Via Tuscolana";
        Record1.Zip = "00181";
        Record1.City = "Roma";
        Record1.CreationDate = new DateOnly(2021, 1, 2);
        Record1.CreationTime = new TimeOnly(12, 0);
        Record1.LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15));

        yield return Record1;

        dynamic Record2 = new ExpandoObject();
        Record2.Name = "Jane";
        Record2.Surname = "Doe";
        Record2.NickName = (string?)"Tarzan lady";
        Record2.BirthDate = new DateOnly(1996, 3, 15);
        Record2.Age = 28;
        Record2.Address = "250 Via Tuscolana";
        Record2.Zip = "00181";
        Record2.City = "Roma";
        Record2.CreationDate = new DateOnly(2023, 5, 25);
        Record2.CreationTime = new TimeOnly(10, 0);
        Record2.LastModifiedDate = new DateTime?(new DateTime(2024, 12, 24, 15, 25, 15));

        yield return Record2;

        await Task.CompletedTask;
    }


    private string GetCellValue(SpreadsheetDocument spreadsheetdocument, Cell cell)
    {
        if (cell.CellValue is not null)
        {
            string value = cell.CellValue.InnerText.Trim();

            if (cell.DataType is not null)
            {
                if (cell.DataType.Value == CellValues.SharedString)
                    return spreadsheetdocument.WorkbookPart?.SharedStringTablePart?.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText ?? value;
                else if (cell.DataType.Value == CellValues.Boolean)
                    return value == "0" ? "FALSE" : "TRUE";
                else if (cell.DataType.Value == CellValues.Date)
                    return DateTime.FromOADate(double.Parse(value)).ToString();
                else
                    return value;
            }
            else
            {
                return value;
            }
        }

        return string.Empty;
    }

    #endregion
}