namespace CsvHelper.OpenXml.Excel;

using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.Abstractions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

/// <summary>
/// The ExcelSaxParser class is defined as sealed, meaning it cannot be inherited. It implements the <see cref="IExcelParser"/> interface, which extends <seealso cref="IParser"/> and IAsyncDisposable.
/// The class provides a robust way to read data from Excel files using the OpenXML SDK, this approach is efficient for reading large Excel files because it doesn't load the entire document into memory and ensures proper resource management.
/// </summary>
public sealed class ExcelSaxParser : IExcelParser
{
    #region Fields

    private readonly OpenXmlHelper OpenXmlHelper = new OpenXmlHelper();

    private string[] CurrentRecord = [];
    private readonly SpreadsheetDocument SpreadsheetDocument;
    private readonly Stream ExcelStream;
    private readonly OpenXmlPartReader PartReader;
    private readonly int LastRow;
    private readonly int LastColumn;

    private int ExcelRow = 0;
    private int ExcelRawRow = 0;

    #endregion

    #region Constructors

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelSaxParser"/> class.
    /// </summary>
    /// <param name="stream">The stream.</param>
    /// <param name="sheetname">The sheet name</param>
    /// <param name="culture">The culture.</param>
    public ExcelSaxParser(Stream stream, string? sheetname, CultureInfo? culture) : this(stream, sheetname, culture is null ? null : new CsvConfiguration(culture)) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelSaxParser"/> class.
    /// </summary>
    /// <param name="stream">The stream.</param>
    /// <param name="sheetname">The sheet name</param>
    /// <param name="configuration">The configuration.</param>
    public ExcelSaxParser(Stream stream, string? sheetname = null, CsvConfiguration? configuration = null)
    {
        SpreadsheetDocument = SpreadsheetDocument.Open(stream, false);

        WorkbookPart WorkbookPart = SpreadsheetDocument.WorkbookPart ?? SpreadsheetDocument.AddWorkbookPart();

        string SheetId = string.IsNullOrEmpty(sheetname) ? WorkbookPart.Workbook.Descendants<Sheet>().First().Id! : WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => sheet.Name == sheetname)?.Id?.Value ?? WorkbookPart.Workbook.Descendants<Sheet>().First().Id!;

        WorksheetPart WorksheetPart = (WorksheetPart)WorkbookPart.GetPartById(SheetId);

        LastRow = WorksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().Count(r => !string.IsNullOrEmpty(r.InnerText)) - 1;

        LastColumn = WorksheetPart.Worksheet.Elements<SheetData>().First().FirstChild?.ChildElements.Count ?? 0;

        Count = WorksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().Select(row => row.Elements<Cell>().Count()).Max();

        PartReader = new OpenXmlPartReader(WorksheetPart);

        Configuration = configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture);

        ExcelStream = stream;

        Context = new CsvContext(this);
    }

    #endregion

    #region Implementation of the IExcelParser interface members

    public string this[int index] => Record?.ElementAtOrDefault(index) ?? string.Empty;

    public long ByteCount => -1;
    public long CharCount => -1;
    public int Count { get; }
    public string[]? Record => CurrentRecord;
    public string RawRecord => Record is null ? string.Empty : string.Join(Delimiter, Record);
    public int Row => ExcelRow;
    public int RawRow => ExcelRawRow;
    public string Delimiter => Configuration.Delimiter;
    public CsvContext Context { get; }
    public IParserConfiguration Configuration { get; }
    public int RowCount { get => LastRow; }

    public string[] GetRecord()
    {
        Row CurrentRow = (Row)PartReader.LoadCurrentElement()!;

        string[] RecordValues;

        int RowCellCount = CurrentRow.Elements<Cell>().Count();

        if (RowCellCount < LastColumn)
        {
            RecordValues = new string[LastColumn];

            for (int i = 0; i < RowCellCount; i++)
            {
                int? ColumnIndex = OpenXmlHelper.GetColumnIndex(CurrentRow.Elements<Cell>().ElementAt(i).CellReference!);

                RecordValues[ColumnIndex!.Value - 1] = GetCellValue(SpreadsheetDocument, CurrentRow.Elements<Cell>().ElementAt(i));
            }
        }
        else
        {
            RecordValues = CurrentRow.Elements<Cell>().Select(c => GetCellValue(SpreadsheetDocument, c)).ToArray();
        }

        return RecordValues;
    }

    public bool Read()
    {
        if (Row > LastRow)
            return false;

        do
        {
            PartReader.Read();

            if (PartReader.ElementType == typeof(Row))
            {
                CurrentRecord = GetRecord();
                ExcelRow++;
                ExcelRawRow++;
            }
        }
        while (PartReader.ElementType != typeof(Row));

        return true;
    }

    public Task<bool> ReadAsync()
    {
        if (Row > LastRow)
            return Task.FromResult(false);

        do
        {
            PartReader.Read();

            if (PartReader.ElementType == typeof(Row))
            {
                CurrentRecord = GetRecord();
                ExcelRow++;
                ExcelRawRow++;
            }
        }
        while (PartReader.ElementType != typeof(Row));

        return Task.FromResult(true);
    }

    #region IDisposable and IAsyncDisposable Methods with Dispose pattern

    private bool Disposed;

    private void Dispose(bool disposing)
    {
        if (!Disposed)
        {
            if (disposing)
            {
                // TODO: dispose managed state (managed objects)
                PartReader.Dispose();
                SpreadsheetDocument.Dispose();
                ExcelStream.Dispose();
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            Disposed = true;
        }
    }

    private async ValueTask DisposeAsyncCore()
    {
        PartReader.Dispose();

        SpreadsheetDocument.Dispose();

        await ExcelStream.DisposeAsync().ConfigureAwait(false);
    }

    public async ValueTask DisposeAsync()
    {
        await DisposeAsyncCore().ConfigureAwait(false);

        Dispose(disposing: false);
        GC.SuppressFinalize(this);
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~ExcelParser()
    // {
    //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
    //     Dispose(disposing: false);
    // }

    #endregion

    #endregion

    #region Private Methods

    private string GetCellValue(SpreadsheetDocument spreadsheetdocument, Cell cell)
    {
        if (cell.CellValue is not null)
        {
            string value = Configuration.TrimOptions.HasFlag(TrimOptions.Trim) ? cell.CellValue.InnerText.Trim() : cell.CellValue.InnerText;

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