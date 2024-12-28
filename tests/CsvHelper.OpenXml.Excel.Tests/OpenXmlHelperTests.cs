namespace CsvHelper.OpenXml.Excel.Tests;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using System;

public class OpenXmlHelperTests
{
    private readonly OpenXmlHelper OpenXmlHelper = new OpenXmlHelper();

    #region Test Methods

    [Fact]
    public void CreateWorksheetStyleShouldCreateStylesheetTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        WorkbookPart WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();

        OpenXmlHelper.CreateWorksheetStyle(ExcelDocument);

        WorkbookStylesPart StylesPart = WorkbookPart.WorkbookStylesPart!;
        StylesPart.Should().NotBeNull();
        StylesPart.Stylesheet.Should().NotBeNull();
    }

    [Fact]
    public void GetSharedStringTablePartShouldReturnSharedStringTablePartTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        WorkbookPart WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();

        SharedStringTablePart SharedStringPart = OpenXmlHelper.GetSharedStringTablePart(WorkbookPart);

        SharedStringPart.Should().NotBeNull();
    }

    [Fact]
    public void InsertWorksheetShouldInsertWorksheetTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        WorkbookPart WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();
        WorkbookPart.Workbook.AppendChild(new Sheets());

        WorksheetPart WorksheetPart = OpenXmlHelper.InsertWorksheet(WorkbookPart, "TestSheet", out var sheetId);

        WorksheetPart.Should().NotBeNull();
        WorksheetPart.Worksheet.Should().NotBeNull();
        sheetId.Should().NotBeNullOrEmpty();
        WorkbookPart.Workbook.Sheets!.Elements<Sheet>().Should().ContainSingle(sheet => sheet.Name == "TestSheet");
    }

    [Fact]
    public void InsertSharedStringItemShouldInsertAndReturnIndexTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();
        SharedStringTablePart SharedStringPart = OpenXmlHelper.GetSharedStringTablePart(WorkbookPart);

        int Index = OpenXmlHelper.InsertSharedStringItem("TestString", SharedStringPart);

        Index.Should().Be(0);
        SharedStringPart.SharedStringTable.Elements<SharedStringItem>().Should().ContainSingle(item => item.InnerText == "TestString");
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(1, "B")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(27, "AB")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    public void GetColumnLettersShouldReturnCorrectLettersTest(int colindex, string expected)
    {
        string ColumnLetterResult = OpenXmlHelper.GetColumnLetters(colindex);

        ColumnLetterResult.Should().Be(expected);
    }

    [Theory]
    [InlineData("A1", 1)]
    [InlineData("B1", 2)]
    [InlineData("Z1", 26)]
    [InlineData("AA1", 27)]
    [InlineData("AB1", 28)]
    [InlineData("AAA1", 703)]
    public void GetColumnIndexShouldReturnCorrectIndexTest(string cellreference, int expected)
    {
        int ColumnIndexResult = OpenXmlHelper.GetColumnIndex(cellreference);

        ColumnIndexResult.Should().Be(expected);
    }

    [Fact]
    public void GetColumnIndexShouldThrowArgumentExceptionForInvalidReferenceTest()
    {
        Action action = () => OpenXmlHelper.GetColumnIndex("1A");

        action.Should().Throw<ArgumentException>().WithMessage("Invalid cell reference format.");
    }

    #endregion
}