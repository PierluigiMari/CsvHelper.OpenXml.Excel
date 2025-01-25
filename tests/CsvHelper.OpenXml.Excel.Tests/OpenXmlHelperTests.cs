namespace CsvHelper.OpenXml.Excel.Tests;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Shouldly;
using System;

public class OpenXmlHelperTests
{
    private readonly OpenXmlHelper OpenXmlHelper = new OpenXmlHelper();

    #region Test Methods

    [Fact]
    public void CreateWorksheetStyle_ShouldCreateStylesheetTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        WorkbookPart WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();

        OpenXmlHelper.CreateWorksheetStyle(ExcelDocument);

        WorkbookStylesPart StylesPart = WorkbookPart.WorkbookStylesPart!;
        StylesPart.ShouldNotBeNull();
        StylesPart.Stylesheet.ShouldNotBeNull();
    }

    [Fact]
    public void GetSharedStringTablePart_ShouldReturnSharedStringTablePartTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        WorkbookPart WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();

        SharedStringTablePart SharedStringPart = OpenXmlHelper.GetSharedStringTablePart(WorkbookPart);

        SharedStringPart.ShouldNotBeNull();
    }

    [Fact]
    public void InsertWorksheet_ShouldInsertWorksheetTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        WorkbookPart WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();
        WorkbookPart.Workbook.AppendChild(new Sheets());

        WorksheetPart WorksheetPart = OpenXmlHelper.InsertWorksheet(WorkbookPart, "TestSheet", out var sheetId);

        WorksheetPart.ShouldNotBeNull();
        WorksheetPart.Worksheet.ShouldNotBeNull();
        sheetId.ShouldNotBeNullOrEmpty();
        WorkbookPart.Workbook.Sheets!.Elements<Sheet>().ShouldHaveSingleItem().Name!.Value.ShouldBe("TestSheet");
    }

    [Fact]
    public void InsertSharedStringItem_ShouldInsertAndReturnIndexTest()
    {
        using MemoryStream ExcelStream = new MemoryStream();
        using SpreadsheetDocument ExcelDocument = SpreadsheetDocument.Create(ExcelStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var WorkbookPart = ExcelDocument.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();
        SharedStringTablePart SharedStringPart = OpenXmlHelper.GetSharedStringTablePart(WorkbookPart);

        Dictionary<string, int> SharedStringDictionary = new Dictionary<string, int>();

        int Index = OpenXmlHelper.InsertSharedStringItem("TestString", SharedStringPart, SharedStringDictionary);

        Index.ShouldBe(0);
        SharedStringPart.SharedStringTable.Elements<SharedStringItem>().ShouldHaveSingleItem().InnerText.ShouldBe("TestString");

        Index = OpenXmlHelper.InsertSharedStringItem("TestString1", SharedStringPart, SharedStringDictionary);
        
        Index.ShouldBe(1);
        SharedStringPart.SharedStringTable.Elements<SharedStringItem>().Count().ShouldBe(2);
        SharedStringPart.SharedStringTable.Elements<SharedStringItem>().ElementAt(1).InnerText.ShouldBe("TestString1");

        Index = OpenXmlHelper.InsertSharedStringItem("TestString", SharedStringPart, SharedStringDictionary);
        Index.ShouldBe(0);
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(1, "B")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(27, "AB")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    [InlineData(16383, "XFD")]
    public void GetColumnLetters_ShouldReturnCorrectLettersTest(int colindex, string expected)
    {
        string ColumnLetterResult = OpenXmlHelper.GetColumnLetters(colindex);

        ColumnLetterResult.ShouldBe(expected);
    }

    [Theory]
    [InlineData("A1", 1)]
    [InlineData("B1", 2)]
    [InlineData("Z1", 26)]
    [InlineData("AA1", 27)]
    [InlineData("AB1", 28)]
    [InlineData("AAA1", 703)]
    [InlineData("XFD1", 16384)]
    public void GetColumnIndex_ShouldReturnCorrectIndexTest(string cellreference, int expected)
    {
        int ColumnIndexResult = OpenXmlHelper.GetColumnIndex(cellreference);

        ColumnIndexResult.ShouldBe(expected);
    }

    [Fact]
    public void GetColumnIndex_ShouldThrowArgumentExceptionForInvalidReferenceTest()
    {
        Action action = () => OpenXmlHelper.GetColumnIndex("1A");

        action.ShouldThrow<ArgumentException>("Invalid cell reference format.");
    }

    #endregion
}