namespace CsvHelper.OpenXml.Excel.Tests;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Shouldly;
using System;
using System.Globalization;
using System.Threading.Tasks;

public class ExcelDomParserTests
{
    #region Test Methods

    [Fact]
    public void Constructor_ShouldInitializeCorrectlyTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.RowCount.ShouldBe(3);
        ExcelParser.Count.ShouldBe(5);
        ExcelParser.Row.ShouldBe(0);
        ExcelParser.RawRow.ShouldBe(0);
    }

    [Fact]
    public void Read_ShouldReturnTrueWhenThereAreMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Read().ShouldBeTrue();
        ExcelParser.Record.ShouldBe(new[] { "Name", "Surname", "NickName", "BirthDate", "Age" });
        ExcelParser.Row.ShouldBe(1);
        ExcelParser.RawRow.ShouldBe(1);

        ExcelParser.Read().ShouldBeTrue();
        ExcelParser.Record.ShouldBe(new[] { "John", "Doe", "", "06/01/1994 00:00:00", "30" });
        ExcelParser.Row.ShouldBe(2);
        ExcelParser.RawRow.ShouldBe(2);

        ExcelParser.Read().ShouldBeTrue();
        ExcelParser.Record.ShouldBe(new[] { "Jane", "Doe", "Tarzan lady", "15/03/1996 00:00:00", "28" });
        ExcelParser.Row.ShouldBe(3);
        ExcelParser.RawRow.ShouldBe(3);
    }

    [Fact]
    public void Read_ShouldReturnFalseWhenNoMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Read().ShouldBeTrue();
        ExcelParser.Read().ShouldBeTrue();
        ExcelParser.Read().ShouldBeTrue();
        ExcelParser.Read().ShouldBeFalse();
    }

    [Fact]
    public async Task ReadAsync_ShouldReturnTrueWhenThereAreMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        (await ExcelParser.ReadAsync()).ShouldBeTrue();
        ExcelParser.Record.ShouldBe(new[] { "Name", "Surname", "NickName", "BirthDate", "Age" });
        ExcelParser.Row.ShouldBe(1);
        ExcelParser.RawRow.ShouldBe(1);
        (await ExcelParser.ReadAsync()).ShouldBeTrue();
        ExcelParser.Record.ShouldBe(new[] { "John", "Doe", "", "06/01/1994 00:00:00", "30" });
        ExcelParser.Row.ShouldBe(2);
        ExcelParser.RawRow.ShouldBe(2);
        (await ExcelParser.ReadAsync()).ShouldBeTrue();
        ExcelParser.Record.ShouldBe(new[] { "Jane", "Doe", "Tarzan lady", "15/03/1996 00:00:00", "28" });
        ExcelParser.Row.ShouldBe(3);
        ExcelParser.RawRow.ShouldBe(3);
    }

    [Fact]
    public async Task ReadAsync_ShouldReturnFalseWhenNoMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        (await ExcelParser.ReadAsync()).ShouldBeTrue();
        (await ExcelParser.ReadAsync()).ShouldBeTrue();
        (await ExcelParser.ReadAsync()).ShouldBeTrue();
        (await ExcelParser.ReadAsync()).ShouldBeFalse();
    }

    [Fact]
    public void Indexer_ShouldReturnCorrectValueTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Read();
        ExcelParser[0].ShouldBe("Name");
        ExcelParser[1].ShouldBe("Surname");
        ExcelParser[2].ShouldBe("NickName");
        ExcelParser[3].ShouldBe("BirthDate");
        ExcelParser[4].ShouldBe("Age");

        ExcelParser.Read();
        ExcelParser[0].ShouldBe("John");
        ExcelParser[1].ShouldBe("Doe");
        ExcelParser[2].ShouldBe("");
        ExcelParser[3].ShouldBe("06/01/1994 00:00:00");
        ExcelParser[4].ShouldBe("30");

        ExcelParser.Read();
        ExcelParser[0].ShouldBe("Jane");
        ExcelParser[1].ShouldBe("Doe");
        ExcelParser[2].ShouldBe("Tarzan lady");
        ExcelParser[3].ShouldBe("15/03/1996 00:00:00");
        ExcelParser[4].ShouldBe("28");
    }

    [Fact]
    public void Dispose_ShouldDisposeResourcesTest()
    {
        MemoryStream ExcelStream = CreateTestExcelStream();
        ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Dispose();

        Action action = () => ExcelParser.Read();
        action.ShouldThrow<ObjectDisposedException>();
    }

    [Fact]
    public async Task DisposeAsync_ShouldDisposeResourcesTest()
    {
        MemoryStream ExcelStream = CreateTestExcelStream();
        ExcelDomParser ExcelParser = new ExcelDomParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        await ExcelParser.DisposeAsync();

        Func<Task> action = async () => await ExcelParser.ReadAsync();
        await action.ShouldThrowAsync<ObjectDisposedException>();
    }

    #endregion

    #region Private Methods

    private MemoryStream CreateTestExcelStream()
    {
        MemoryStream ExcelStream = new MemoryStream();

        using (SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart WorkbookPart = ExcelDocument.AddWorkbookPart();
            WorkbookPart.Workbook = new Workbook();
            WorksheetPart WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            WorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets Sheets = ExcelDocument.WorkbookPart!.Workbook.AppendChild(new Sheets());
            Sheet Sheet = new Sheet() { Id = ExcelDocument.WorkbookPart.GetIdOfPart(WorksheetPart), SheetId = 1, Name = "Sheet1" };
            Sheets.Append(Sheet);

            SheetData SheetData = WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            Row Row = new Row();

            Row.Append(new Cell() { CellValue = new CellValue("Name"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue("Surname"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue("NickName"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue("BirthDate"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue("Age"), DataType = CellValues.String });
            SheetData.Append(Row);

            Row = new Row();
            Row.Append(new Cell() { CellValue = new CellValue("John"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue("Doe"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue(""), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue(new DateTime(1994, 1, 6).ToOADate()), DataType = CellValues.Date });
            Row.Append(new Cell() { CellValue = new CellValue("30"), DataType = CellValues.Number });
            SheetData.Append(Row);

            Row = new Row();
            Row.Append(new Cell() { CellValue = new CellValue("Jane"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue("Doe"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue("Tarzan lady"), DataType = CellValues.String });
            Row.Append(new Cell() { CellValue = new CellValue(new DateTime(1996, 3, 15).ToOADate()), DataType = CellValues.Date });
            Row.Append(new Cell() { CellValue = new CellValue("28"), DataType = CellValues.Number });
            SheetData.Append(Row);
        }

        ExcelStream.Position = 0;

        return ExcelStream;
    }

    #endregion
}