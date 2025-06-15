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
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

/// <summary>
/// The ExcelSaxParser class is defined as sealed, meaning it cannot be inherited. It implements the <see cref="IExcelParser"/> interface, which extends <seealso cref="IParser"/> and IAsyncDisposable.
/// The class provides a robust way to read data from Excel files using the OpenXML SDK, this approach is efficient for reading large Excel files because it doesn't load the entire document into memory and ensures proper resource management.
/// </summary>
public sealed class ExcelSaxParser : IExcelParser
{
    #region Fields

    /// <summary>
    /// Helper class for OpenXML operations.
    /// </summary>
    private readonly OpenXmlHelper OpenXmlHelper = new OpenXmlHelper();

    /// <summary>
    /// The current record being processed as an array of strings.
    /// </summary>
    private string[] CurrentRecord = [];
    /// <summary>
    /// The SpreadsheetDocument instance representing the Excel file.
    /// </summary>
    private readonly SpreadsheetDocument SpreadsheetDocument;
    /// <summary>
    /// The stream containing the Excel file data.
    /// </summary>
    private readonly Stream ExcelStream;
    /// <summary>
    /// The relationship id of sheet.
    /// </summary>
    private readonly string SheetId;
    /// <summary>
    /// The OpenXmlPartReader instance for reading parts of the Excel file.
    /// </summary>
    private readonly OpenXmlPartReader PartReader;
    /// <summary>
    /// The index of the last row in the Excel sheet.
    /// </summary>
    private readonly int LastRow;
    /// <summary>
    /// The index of the last column in the Excel sheet.
    /// </summary>
    private readonly int LastColumn;

    /// <summary>
    /// The current row index being processed.
    /// </summary>
    private int ExcelRow = 0;
    /// <summary>
    /// The raw row index being processed.
    /// </summary>
    private int ExcelRawRow = 0;

    /// <summary>
    /// Represents a collection of hyperlinks.
    /// </summary>
    private IEnumerable<Hyperlink>? Hyperlinks = null;

    #endregion

    #region Constructors

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelSaxParser"/> class with the specified stream, sheet name, and culture.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data.</param>
    /// <param name="sheetname">The name of the sheet to parse. If <c>null</c> or not specified, the first sheet in the workbook is used. The sheet name must not exceed 31 characters (as per Excel's limitations), otherwise, an <see cref="ArgumentOutOfRangeException"/> is thrown.</param>
    /// <param name="culture">The culture information for parsing.</param>
    public ExcelSaxParser(Stream stream, string? sheetname, CultureInfo? culture) : this(stream, sheetname, culture is null ? null : new CsvConfiguration(culture)) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelSaxParser"/> class with the specified stream, sheet name, and configuration.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data.</param>
    /// <param name="sheetname">The name of the sheet to parse. If <c>null</c> or not specified, the first sheet in the workbook is used. The sheet name must not exceed 31 characters (as per Excel's limitations), otherwise, an <see cref="ArgumentOutOfRangeException"/> is thrown.</param>
    /// <param name="configuration">The CSV configuration for parsing.</param>
    public ExcelSaxParser(Stream stream, string? sheetname = null, CsvConfiguration? configuration = null)
    {
        ArgumentNullException.ThrowIfNull(stream, nameof(stream));

        if (sheetname is not null)
            ArgumentOutOfRangeException.ThrowIfGreaterThan(sheetname.Length, 31, nameof(sheetname.Length));

        SpreadsheetDocument = SpreadsheetDocument.Open(stream, false);

        WorkbookPart WorkbookPart = SpreadsheetDocument.WorkbookPart ?? SpreadsheetDocument.AddWorkbookPart();

        SheetId = string.IsNullOrEmpty(sheetname) ? WorkbookPart.Workbook.Descendants<Sheet>().First().Id! : WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => sheet.Name == sheetname)?.Id?.Value ?? WorkbookPart.Workbook.Descendants<Sheet>().First().Id!;

        WorksheetPart WorksheetPart = (WorksheetPart)WorkbookPart.GetPartById(SheetId);

        LastRow = WorksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().Count(r => !string.IsNullOrEmpty(r.InnerText)) - 1;

        LastColumn = WorksheetPart.Worksheet.Elements<SheetData>().First().FirstChild?.ChildElements.Count ?? 0;

        Count = WorksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().Select(row => row.Elements<Cell>().Count()).Max();

        PartReader = new OpenXmlPartReader(WorksheetPart);

        Configuration = configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture);

        ExcelStream = stream;

        Context = new CsvContext(this);

        GetHyperlinks();
    }

    #endregion

    #region Implementation of the IExcelParser interface members

    /// <summary>
    /// Gets the value of the cell at the specified index in the current record.
    /// </summary>
    /// <param name="index">The index of the cell.</param>
    /// <returns>The value of the cell.</returns>
    public string this[int index] => Record?.ElementAtOrDefault(index) ?? string.Empty;

    /// <summary>
    /// Gets the total number of bytes read. Not applicable in ExcelParser, not used. Always returns -1.
    /// </summary>
    public long ByteCount => -1;
    /// <summary>
    /// Gets the total number of characters read. Not applicable in ExcelParser, not used. Always returns -1.
    /// </summary>
    public long CharCount => -1;
    /// <summary>
    /// Gets the total number of columns in the current record.
    /// </summary>
    public int Count { get; }
    /// <summary>
    /// Gets the current record as an array of strings.
    /// </summary>
    public string[]? Record => CurrentRecord;
    /// <summary>
    /// Gets the raw record as a single string.
    /// </summary>
    public string RawRecord => Record is null ? string.Empty : string.Join(Delimiter, Record);
    /// <summary>
    /// Gets the current row index being processed.
    /// </summary>
    public int Row => ExcelRow;
    /// <summary>
    /// Gets the raw row index being processed.
    /// </summary>
    public int RawRow => ExcelRawRow;
    /// <summary>
    /// Gets the delimiter used to separate fields. The default value is <see cref="TextInfo.ListSeparator"/>. Not applicable in ExcelParser, not used.
    /// </summary>
    public string Delimiter => Configuration.Delimiter;
    /// <summary>
    /// Gets the reading context of parser.
    /// </summary>
    public CsvContext Context { get; }
    /// <summary>
    /// Gets the configuration of parser.
    /// </summary>
    public IParserConfiguration Configuration { get; }
    /// <summary>
    /// Gets the total number of rows in the Excel file.
    /// </summary>
    public int RowCount { get => LastRow + 1; }

    /// <summary>
    /// Retrieves the record as an array of strings.
    /// </summary>
    /// <returns>An array of strings representing the record.</returns>
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

    /// <summary>
    /// Reads the next record from the Excel sheet.
    /// </summary>
    /// <returns>True if the next record was read successfully; otherwise, false.</returns>
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

    /// <summary>
    /// Asynchronously reads the next record from the Excel sheet.
    /// </summary>
    /// <returns>A task that represents the asynchronous read operation. The task result contains true if the next record was read successfully; otherwise, false.</returns>
    public Task<bool> ReadAsync() => Task.FromResult(Read());

    #region IDisposable and IAsyncDisposable Methods with Dispose pattern

    /// <summary>
    /// Indicates whether the object has been disposed.
    /// </summary>
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

    /// <inheritdoc/>
    public async ValueTask DisposeAsync()
    {
        await DisposeAsyncCore().ConfigureAwait(false);

        Dispose(disposing: false);
        GC.SuppressFinalize(this);
    }

    /// <inheritdoc/>
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

    /// <summary>
    /// Gets the value of the specified cell.
    /// </summary>
    /// <param name="spreadsheetdocument">The spreadsheet document instance containing the cell.</param>
    /// <param name="cell">The cell to retrieve the value from.</param>
    /// <returns>The value of the cell as a string.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private string GetCellValue(SpreadsheetDocument spreadsheetdocument, Cell cell)
    {
        if (cell.CellValue is null)
        {
            return string.Empty;
        }

        string value = Configuration.TrimOptions.HasFlag(TrimOptions.Trim) ? cell.CellValue.InnerText.Trim() : cell.CellValue.InnerText;

        if(Hyperlinks is null)
        {
            if (cell.DataType is null)
            {
                return value;
            }

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
            Hyperlink? Hyperlink = Hyperlinks?.FirstOrDefault(h => h.Reference == cell.CellReference);

            if (Hyperlink is null)
            {
                if (cell.DataType is null)
                {
                    return value;
                }

                if (cell.DataType.Value == CellValues.SharedString)
                    return spreadsheetdocument.WorkbookPart?.SharedStringTablePart?.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText ?? value;
                else if (cell.DataType.Value == CellValues.Boolean)
                    return value == "0" ? "FALSE" : "TRUE";
                else if (cell.DataType.Value == CellValues.Date)
                    return DateTime.FromOADate(double.Parse(value)).ToString();
                else
                    return value;
            }

            const string Delimiter = "(|->)";

            if (Hyperlink.Id is not null)
            {
                string? HyperlinkAbsoluteUri = ((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)).HyperlinkRelationships.FirstOrDefault(hr => hr.Id == Hyperlink.Id)?.Uri?.AbsoluteUri;

                if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
                {
                    string CellValue = spreadsheetdocument.WorkbookPart?.SharedStringTablePart?.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText ?? value;

                    return CellValue.Equals(HyperlinkAbsoluteUri, StringComparison.OrdinalIgnoreCase) ? CellValue : string.Concat(CellValue, Delimiter, HyperlinkAbsoluteUri);
                }

                return value.Equals(HyperlinkAbsoluteUri, StringComparison.OrdinalIgnoreCase) ? value : string.Concat(value, Delimiter, HyperlinkAbsoluteUri);

            }

            if (Hyperlink.Location is not null)
            {
                if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
                {
                    string CellValue = spreadsheetdocument.WorkbookPart?.SharedStringTablePart?.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText ?? value;

                    return CellValue.Equals(Hyperlink.Location, StringComparison.OrdinalIgnoreCase) ? CellValue : string.Concat(CellValue, Delimiter, Hyperlink.Location);
                }

                return value.Equals(Hyperlink.Location, StringComparison.OrdinalIgnoreCase) ? value : string.Concat(value, Delimiter, Hyperlink.Location);
            }

            return value;
        }
    }

    /// <summary>
    /// Extracts all hyperlinks from the worksheet associated with the current sheet ID.
    /// </summary>
    private void GetHyperlinks()
    {
        using OpenXmlPartReader WorksheetPartReader = new OpenXmlPartReader(((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)));

        do
        {
            WorksheetPartReader.Read();

            if (WorksheetPartReader.ElementType == typeof(Worksheet))
            {
                Worksheet Worksheet = ((Worksheet)WorksheetPartReader.LoadCurrentElement()!);

                Hyperlink[] ArrayHyperlinks = Worksheet.Descendants<Hyperlink>().ToArray();

                Hyperlinks = ArrayHyperlinks.Length > 0 ? ArrayHyperlinks : null;
            }
        }
        while (WorksheetPartReader.ElementType != typeof(Worksheet));
    }

    #endregion
}