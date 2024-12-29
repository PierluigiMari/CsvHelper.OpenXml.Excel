namespace CsvHelper.OpenXml.Excel.Tests;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using System;
using System.Globalization;
using System.Threading.Tasks;

public class ExcelSaxParserTests
{
    #region Test Methods

    [Fact]
    public void Constructor_ShouldInitializeCorrectlyTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.RowCount.Should().Be(3);
        ExcelParser.Count.Should().Be(5);
        ExcelParser.Row.Should().Be(0);
        ExcelParser.RawRow.Should().Be(0);
    }

    [Fact]
    public void Read_ShouldReturnTrueWhenThereAreMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Read().Should().BeTrue();
        ExcelParser.Record.Should().BeEquivalentTo(new[] { "Name", "Surname", "NickName", "BirthDate", "Age" });
        ExcelParser.Row.Should().Be(1);
        ExcelParser.RawRow.Should().Be(1);

        ExcelParser.Read().Should().BeTrue();
        ExcelParser.Record.Should().BeEquivalentTo(new[] { "John", "Doe", "", "06/01/1994 00:00:00", "30" });
        ExcelParser.Row.Should().Be(2);
        ExcelParser.RawRow.Should().Be(2);

        ExcelParser.Read().Should().BeTrue();
        ExcelParser.Record.Should().BeEquivalentTo(new[] { "Jane", "Doe", "Tarzan lady", "15/03/1996 00:00:00", "28" });
        ExcelParser.Row.Should().Be(3);
        ExcelParser.RawRow.Should().Be(3);
    }

    [Fact]
    public void Read_ShouldReturnFalseWhenNoMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Read().Should().BeTrue();
        ExcelParser.Read().Should().BeTrue();
        ExcelParser.Read().Should().BeTrue();
        ExcelParser.Read().Should().BeFalse();
    }

    [Fact]
    public async Task ReadAsync_ShouldReturnTrueWhenThereAreMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        (await ExcelParser.ReadAsync()).Should().BeTrue();
        ExcelParser.Record.Should().BeEquivalentTo(new[] { "Name", "Surname", "NickName", "BirthDate", "Age" });
        ExcelParser.Row.Should().Be(1);
        ExcelParser.RawRow.Should().Be(1);
        (await ExcelParser.ReadAsync()).Should().BeTrue();
        ExcelParser.Record.Should().BeEquivalentTo(new[] { "John", "Doe", "", "06/01/1994 00:00:00", "30" });
        ExcelParser.Row.Should().Be(2);
        ExcelParser.RawRow.Should().Be(2);
        (await ExcelParser.ReadAsync()).Should().BeTrue();
        ExcelParser.Record.Should().BeEquivalentTo(new[] { "Jane", "Doe", "Tarzan lady", "15/03/1996 00:00:00", "28" });
        ExcelParser.Row.Should().Be(3);
        ExcelParser.RawRow.Should().Be(3);
    }

    [Fact]
    public async Task ReadAsync_ShouldReturnFalseWhenNoMoreRowsTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        (await ExcelParser.ReadAsync()).Should().BeTrue();
        (await ExcelParser.ReadAsync()).Should().BeTrue();
        (await ExcelParser.ReadAsync()).Should().BeTrue();
        (await ExcelParser.ReadAsync()).Should().BeFalse();
    }

    [Fact]
    public void Indexer_ShouldReturnCorrectValueTest()
    {
        using MemoryStream ExcelStream = CreateTestExcelStream();
        using ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Read();
        ExcelParser[0].Should().Be("Name");
        ExcelParser[1].Should().Be("Surname");
        ExcelParser[2].Should().Be("NickName");
        ExcelParser[3].Should().Be("BirthDate");
        ExcelParser[4].Should().Be("Age");

        ExcelParser.Read();
        ExcelParser[0].Should().Be("John");
        ExcelParser[1].Should().Be("Doe");
        ExcelParser[2].Should().Be("");
        ExcelParser[3].Should().Be("06/01/1994 00:00:00");
        ExcelParser[4].Should().Be("30");

        ExcelParser.Read();
        ExcelParser[0].Should().Be("Jane");
        ExcelParser[1].Should().Be("Doe");
        ExcelParser[2].Should().Be("Tarzan lady");
        ExcelParser[3].Should().Be("15/03/1996 00:00:00");
        ExcelParser[4].Should().Be("28");
    }

    [Fact]
    public void Dispose_ShouldDisposeResourcesTest()
    {
        MemoryStream ExcelStream = CreateTestExcelStream();
        ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        ExcelParser.Dispose();

        Action action = () => ExcelParser.Read();
        action.Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public async Task DisposeAsync_ShouldDisposeResourcesTest()
    {
        MemoryStream ExcelStream = CreateTestExcelStream();
        ExcelSaxParser ExcelParser = new ExcelSaxParser(ExcelStream, "Sheet1", CultureInfo.InvariantCulture);

        await ExcelParser.DisposeAsync();

        Func<Task> action = async () => await ExcelParser.ReadAsync();
        await action.Should().ThrowAsync<ObjectDisposedException>();
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