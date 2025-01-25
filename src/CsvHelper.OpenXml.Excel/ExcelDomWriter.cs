namespace CsvHelper.OpenXml.Excel;

using CsvHelper.Configuration;
using CsvHelper.OpenXml.Excel.Abstractions;
using CsvHelper.OpenXml.Excel.TypeConversion;
using CsvHelper.TypeConversion;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// The ExcelDomWriter class is defined as sealed, meaning it cannot be inherited. It implements the <see cref="IExcelWriter"/> interface, which extends <seealso cref="IWriter"/>.
/// The ExcelDomWriter class provides a robust way to write data to Excel files using the OpenXML SDK, extending the capabilities of <see cref="CsvWriter"/> to support Excel specific features and formats. It handles various data types, manages the Excel document structure, and ensures proper resource management.
/// </summary>
public sealed class ExcelDomWriter : CsvWriter, IExcelWriter
{
    #region Fields

    /// <summary>
    /// Helper class for OpenXML operations.
    /// </summary>
    private readonly OpenXmlHelper OpenXmlHelper = new OpenXmlHelper();

    /// <summary>
    /// Options for handling injection attacks.
    /// </summary>
    private readonly InjectionOptions InjectionOptions;

    /// <summary>
    /// The SpreadsheetDocument being written to.
    /// </summary>
    private readonly SpreadsheetDocument SpreadsheetDocument;

    /// <summary>
    /// The SharedStringTablePart of the SpreadsheetDocument.
    /// </summary>
    private readonly SharedStringTablePart SharedStringPart;

    /// <summary>
    /// The WorksheetPart currently being written to.
    /// </summary>
    private WorksheetPart WorksheetPart = null!;

    /// <summary>
    /// The relationship id of current sheet.
    /// </summary>
    private string SheetId = string.Empty;

    /// <summary>
    /// The current row index in the Excel sheet.
    /// </summary>
    private int ExcelRowIndex = 1;
    /// <summary>
    /// The current column index in the Excel sheet.
    /// </summary>
    private int ExcelColumnIndex = 0;
    /// <summary>
    /// The total number of columns in the Excel sheet.
    /// </summary>
    private int ExcelColumnCount = 0;

    /// <summary>
    /// The current row being written to.
    /// </summary>
    private Row WritingRow = null!;

    /// <summary>
    /// The type of the field currently being written.
    /// </summary>
    private Type? WritingFieldType = null!;

    /// <summary>
    /// Details of the Excel cell member map.
    /// </summary>
    private readonly Dictionary<int, (string FieldTypeName, ExcelCellFormats? ExcelCellFormat, double CellLength)> ExcelCellMemberMapDetails = new Dictionary<int, (string FieldTypeName, ExcelCellFormats? ExcelCellFormat, double CellLength)>();

    /// <summary>
    /// The name of the last sheet written to.
    /// </summary>
    private string? LastSheetName;

    /// <summary>
    /// Dictionary to store shared string items and their corresponding indices.
    /// </summary>
    private Dictionary<string, int> SharedStringDictionary = new Dictionary<string, int>();

    #endregion

    #region Constructors

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelDomWriter"/> class with the specified stream and culture.
    /// </summary>
    /// <param name="stream">The stream to write to.</param>
    /// <param name="culture">The culture to use for formatting.</param>
    public ExcelDomWriter(Stream stream, CultureInfo? culture) : this(stream, culture is null ? new CsvConfiguration(CultureInfo.InvariantCulture) : new CsvConfiguration(culture)) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelDomWriter"/> class with the specified stream and configuration.
    /// </summary>
    /// <param name="stream">The stream to write to.</param>
    /// <param name="configuration">The configuration to use.</param>
    public ExcelDomWriter(Stream stream, CsvConfiguration? configuration = null) : base(TextWriter.Null, configuration is null ? new CsvConfiguration(CultureInfo.InvariantCulture) : configuration)
    {
        base.Configuration.Validate();

        if (stream.Length > 0)
            SpreadsheetDocument = SpreadsheetDocument.Open(stream, true);
        else
            SpreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart WorkbookPart = SpreadsheetDocument.WorkbookPart ?? SpreadsheetDocument.AddWorkbookPart();

        WorkbookPart.Workbook ??= new Workbook();

        OpenXmlHelper.CreateWorksheetStyle(SpreadsheetDocument);

        SharedStringPart = OpenXmlHelper.GetSharedStringTablePart(WorkbookPart);

        InjectionOptions = base.Configuration.InjectionOptions;
    }

    #endregion

    #region Overriding CsvWriter Properties and Methods

    /// <summary>
    /// Gets the current row index.
    /// </summary>
    public override int Row => ExcelRowIndex;
    /// <summary>
    /// Gets the current column index.
    /// </summary>
    public override int Index => ExcelColumnIndex;


    /// <summary>
    /// Writes a field to the current row.
    /// </summary>
    /// <typeparam name="T">The type of the field.</typeparam>
    /// <param name="field">The field to write.</param>
    /// <param name="converter">The converter to use for the field.</param>
    public override void WriteField<T>(T? field, ITypeConverter converter) where T : default
    {
        WritingFieldType = typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(typeof(T)) : typeof(T);

        base.WriteField(field, converter);
    }

    /// <summary>
    /// Writes a converted field to the current row.
    /// </summary>
    /// <param name="field">The field to write.</param>
    /// <param name="fieldType">The type of the field.</param>
    public override void WriteConvertedField(string? field, Type fieldType)
    {
        WritingFieldType = fieldType.IsGenericType && fieldType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(fieldType) : fieldType;

        base.WriteConvertedField(field, fieldType);
    }

    /// <summary>
    /// Writes a field to the current row.
    /// </summary>
    /// <param name="field">The field to write.</param>
    /// <param name="shouldQuote">Whether the field should be quoted.</param>
    public override void WriteField(string? field, bool shouldQuote)
    {
        if (InjectionOptions == InjectionOptions.Strip)
        {
            field = SanitizeForInjection(field);
        }

        WriteToCell(field);

        ExcelColumnIndex++;
    }


    /// <summary>
    /// Moves to the next record (row).
    /// </summary>
    public override void NextRecord()
    {
        ExcelColumnCount = ExcelColumnIndex;
        ExcelColumnIndex = 0;
        ExcelRowIndex++;

        WritingRow = new Row();
        ((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)).Worksheet.Elements<SheetData>().First().Append(WritingRow);
    }

    /// <summary>
    /// Asynchronously moves to the next record (row).
    /// </summary>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public override Task NextRecordAsync()
    {
        NextRecord();

        return Task.CompletedTask;
    }

    /// <summary>
    /// Flushes the current data to the Excel document.
    /// </summary>
    public override void Flush()
    {
        SpreadsheetDocument.WorkbookPart!.Workbook.Save();
    }

    /// <summary>
    /// Asynchronously flushes the current data to the Excel document.
    /// </summary>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public override Task FlushAsync()
    {
        SpreadsheetDocument.WorkbookPart!.Workbook.Save();

        return Task.CompletedTask;
    }

    #region IDisposable and IAsyncDisposable override Methods of CsvWriter

    /// <summary>
    /// Indicates whether the object has been disposed.
    /// </summary>
    private bool Disposed;

    /// <inheritdoc/>
    protected override void Dispose(bool disposing)
    {
        if (Disposed)
            return;

        SpreadsheetDocument.WorkbookPart!.Workbook.Save();

        if (disposing)
        {
            SpreadsheetDocument.Dispose();
        }

        // TODO: free unmanaged resources (unmanaged objects) and override finalizer
        // TODO: set large fields to null
        Disposed = true;
    }

    /// <inheritdoc/>
    protected override ValueTask DisposeAsync(bool disposing)
    {
        if (Disposed)
            return ValueTask.CompletedTask;

        SpreadsheetDocument.WorkbookPart!.Workbook.Save();

        if (disposing)
        {
            SpreadsheetDocument.Dispose();
        }

        // TODO: free unmanaged resources (unmanaged objects) and override finalizer
        // TODO: set large fields to null
        Disposed = true;

        return ValueTask.CompletedTask;
    }

    #endregion

    #endregion

    #region Implementation of the IExcelWriter interface members

    /// <summary>
    /// Writes a single record to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    /// <param name="record">The record to write.</param>
    /// <param name="sheetname">The name of the sheet to write to. If null, the default sheet is used.</param>
    public void WriteRecord<T>(T? record, string? sheetname = null)
    {
        if (ExcelRowIndex == 1)
        {
            InitializeWritingNewWorksheet<T>(sheetname);

            if (typeof(T) == typeof(ExpandoObject))
                WriteDynamicHeader(record as ExpandoObject);
            else
                WriteHeader<T>();

            NextRecord();
        }

        base.WriteRecord(record);
    }

    /// <summary>
    /// Writes multiple records to the specified sheet.
    /// </summary>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to. If null, the default sheet is used.</param>
    public void WriteRecords(IEnumerable records, string? sheetname = null)
    {
        IEnumerator Enumerator = records.GetEnumerator();

        Type type;
        if (Enumerator.MoveNext())
        {
            type = Enumerator.Current.GetType();
            Enumerator.Reset();
        }
        else
        {
            return;
        }

        if (ExcelRowIndex == 1)
        {
            InitializeWritingNewWorksheet(type, sheetname);
        }
        else
        {
            if (LastSheetName != sheetname)
            {
                ExcelRowIndex = 1;

                ExcelCellMemberMapDetails.Clear();

                InitializeWritingNewWorksheet(type, sheetname);

                WriteHeader(type);
                NextRecord();
            }
        }

        LastSheetName = sheetname;

        base.WriteRecords(records);

        ((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)).Worksheet.Elements<SheetData>().First().RemoveChild(WritingRow);

        ExcelRowIndex--;

        AutoFitColumns();
    }

    /// <summary>
    /// Asynchronously writes multiple records to the specified sheet.
    /// </summary>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to. If null, the default sheet is used.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    public async Task WriteRecordsAsync(IEnumerable records, string? sheetname = null, CancellationToken cancellationToken = default)
    {
        IEnumerator Enumerator = records.GetEnumerator();

        Type type;
        if (Enumerator.MoveNext())
        {
            type = Enumerator.Current.GetType();
            Enumerator.Reset();
        }
        else
        {
            return;
        }

        if (ExcelRowIndex == 1)
        {
            InitializeWritingNewWorksheet(type, sheetname);
        }
        else
        {
            if (LastSheetName != sheetname)
            {
                ExcelRowIndex = 1;

                ExcelCellMemberMapDetails.Clear();

                InitializeWritingNewWorksheet(type, sheetname);

                WriteHeader(type);
                NextRecord();
            }
        }


        LastSheetName = sheetname;

        await base.WriteRecordsAsync(records, cancellationToken);

        ((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)).Worksheet.Elements<SheetData>().First().RemoveChild(WritingRow);

        ExcelRowIndex--;

        AutoFitColumns();
    }


    /// <summary>
    /// Writes multiple records of a specific type to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the records.</typeparam>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to.</param>
    public void WriteRecords<T>(IEnumerable<T> records, string? sheetname = null)
    {
        if (ExcelRowIndex == 1)
        {
            InitializeWritingNewWorksheet<T>(sheetname);
        }
        else
        {
            if (LastSheetName != sheetname)
            {
                ExcelRowIndex = 1;

                ExcelCellMemberMapDetails.Clear();

                InitializeWritingNewWorksheet<T>(sheetname);

                WriteHeader<T>();
                NextRecord();
            }
        }


        LastSheetName = sheetname;

        base.WriteRecords(records);

        ((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)).Worksheet.Elements<SheetData>().First().RemoveChild(WritingRow);

        ExcelRowIndex--;

        AutoFitColumns();
    }

    /// <summary>
    /// Asynchronously writes multiple records of a specific type to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the records.</typeparam>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    public async Task WriteRecordsAsync<T>(IEnumerable<T> records, string? sheetname = null, CancellationToken cancellationToken = default)
    {
        if (ExcelRowIndex == 1)
        {
            InitializeWritingNewWorksheet<T>(sheetname);
        }
        else
        {
            if (LastSheetName != sheetname)
            {
                ExcelRowIndex = 1;

                ExcelCellMemberMapDetails.Clear();

                InitializeWritingNewWorksheet<T>(sheetname);

                WriteHeader<T>();
                NextRecord();
            }
        }


        LastSheetName = sheetname;

        await base.WriteRecordsAsync(records, cancellationToken);

        ((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)).Worksheet.Elements<SheetData>().First().RemoveChild(WritingRow);

        ExcelRowIndex--;

        AutoFitColumns();
    }


    /// <summary>
    /// Asynchronously writes multiple records from an asynchronous enumerable to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the records.</typeparam>
    /// <param name="records">The asynchronous enumerable of records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    public async Task WriteRecordsAsync<T>(IAsyncEnumerable<T> records, string? sheetname = null, CancellationToken cancellationToken = default)
    {
        if (ExcelRowIndex == 1)
        {
            InitializeWritingNewWorksheet<T>(sheetname);
        }
        else
        {
            if (LastSheetName != sheetname)
            {
                ExcelRowIndex = 1;

                ExcelCellMemberMapDetails.Clear();

                InitializeWritingNewWorksheet<T>(sheetname);

                WriteHeader<T>();
                NextRecord();
            }
        }


        LastSheetName = sheetname;

        await base.WriteRecordsAsync(records, cancellationToken);

        ((WorksheetPart)SpreadsheetDocument.WorkbookPart!.GetPartById(SheetId)).Worksheet.Elements<SheetData>().First().RemoveChild(WritingRow);

        ExcelRowIndex--;

        AutoFitColumns();
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Initializes a new worksheet for writing.
    /// </summary>
    /// <param name="type">The type of the records.</param>
    /// <param name="sheetname">The name of the sheet.</param>
    private void InitializeWritingNewWorksheet(Type type, string? sheetname)
    {
        WorksheetPart = OpenXmlHelper.InsertWorksheet(SpreadsheetDocument.WorkbookPart!, string.IsNullOrEmpty(sheetname) ? null : sheetname, out SheetId);

        SheetData SheetData = WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;

        ClassMap? ClassMap = Context.Maps[type];
        if (ClassMap is not null)
        {
            IEnumerable<MemberMapData> MemberMapData = ClassMap.MemberMaps.Select(x => x.Data);

            foreach (MemberMapData MemberMapDataItem in MemberMapData)
            {
                ExcelCellMemberMapDetails.Add(MemberMapDataItem.Index, (MemberMapDataItem.Type.Name, MemberMapDataItem.TypeConverterOptions is ExcelTypeConverterOptions ExcelTypeConverterOption ? ExcelTypeConverterOption.ExcelCellFormat : null, 0));
            }
        }

        WritingRow = new Row();
        SheetData.Append(WritingRow);
    }

    /// <summary>
    /// Initializes a new worksheet for writing.
    /// </summary>
    /// <typeparam name="T">The type of the records.</typeparam>
    /// <param name="sheetname">The name of the sheet.</param>
    private void InitializeWritingNewWorksheet<T>(string? sheetname)
    {
        InitializeWritingNewWorksheet(typeof(T), sheetname);
    }

    /// <summary>
    /// Writes a value to a cell.
    /// </summary>
    /// <param name="value">The value to write.</param>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void WriteToCell(string? value)
    {
        if (value is null || value.Length == 0)
        {
            WritingFieldType = null;
            return;
        }

        if (ExcelCellMemberMapDetails.Count > 0 && ExcelCellMemberMapDetails[ExcelColumnIndex].CellLength < value.Length)
            ExcelCellMemberMapDetails[ExcelColumnIndex] = (ExcelCellMemberMapDetails[ExcelColumnIndex].FieldTypeName, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat, value.Length);

        Cell Cell = new Cell() { CellReference = $"{OpenXmlHelper.GetColumnLetters(ExcelColumnIndex)}{ExcelRowIndex}" };
        WritingRow.Append(Cell);

        if (WritingFieldType is null)
        {
            WriteStringOrGuidToCellWithUseSharedStrings(value, Cell, null, true);
        }
        else
        {
            WriteSpecificTypeInCell(value, Cell);
        }

        WritingFieldType = null;
    }

    /// <summary>
    /// Writes a specific type of value to a cell.
    /// </summary>
    /// <param name="value">The value to write.</param>
    /// <param name="cell">The cell to write the value to.</param>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void WriteSpecificTypeInCell(string value, Cell cell)
    {
        //ExcelCellFormats? ExcelCellFormat =ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat;
        ExcelCellFormats? ExcelCellFormat = ExcelCellMemberMapDetails.Count > 0 ? ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat : null;


        switch (Type.GetTypeCode(WritingFieldType))
        {
            case TypeCode.String:
            case TypeCode.Object when WritingFieldType == typeof(Guid):
                WriteStringOrGuidToCellWithUseSharedStrings(value, cell, ExcelCellFormat);
                break;
            case TypeCode.DateTime:
                WriteDateTimeToCellWithUseSharedStrings(value, cell, ExcelCellFormat);
                break;
            case TypeCode.Int32:
                WriteIntToCell(value, cell);
                break;
            case TypeCode.Decimal:
                WriteDecimalToCell(value, cell, ExcelCellFormat);
                break;
            case TypeCode.Double:
                WriteDoubleToCell(value, cell, ExcelCellFormat);
                break;
            case TypeCode.Boolean:
                WriteBoolToCell(value, cell);
                break;
            default:
                if (WritingFieldType == typeof(DateOnly))
                    WriteDateOnlyToCellWithUseSharedStrings(value, cell, ExcelCellFormat);
                else if (WritingFieldType == typeof(TimeOnly))
                    WriteTimeOnlyToCellWithUseSharedStrings(value, cell, ExcelCellFormat);
                else
                    throw new NotImplementedException($"Writing of the specific type {WritingFieldType!.Name} not yet implemented!");
                break;
        }
    }

    /// <summary>
    /// Writes a string or GUID value to a cell.
    /// </summary>
    /// <param name="value">The value to write.</param>
    /// <param name="cell">The cell to write the value to.</param>
    /// <param name="excelcellformat">The format to apply to the cell.</param>
    /// <param name="isheadercell">Indicates whether the cell is a header cell.</param>
    //[MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void WriteStringOrGuidToCellWithUseSharedStrings(string value, Cell cell, ExcelCellFormats? excelcellformat, bool isheadercell = false)
    {
        int index;
        if (int.TryParse(value, out _))
        {
            if (excelcellformat is not null && excelcellformat == ExcelCellFormats.Text)
            {
                index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart, SharedStringDictionary);

                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = (uint)excelcellformat;
            }
            else
            {
                cell.CellValue = new CellValue(value);

                if (excelcellformat is not null)
                    cell.StyleIndex = (uint)excelcellformat;
            }
        }
        else
        {
            index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart, SharedStringDictionary);

            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            if (isheadercell)
                cell.StyleIndex = (uint)ExcelCellFormats.DefaultBoldCentered;
        }
    }

    /// <summary>
    /// Writes a DateOnly value to a cell.
    /// </summary>
    /// <param name="value">The DateOnly value to write.</param>
    /// <param name="cell">The cell to write the value to.</param>
    /// <param name="excelcellformat">The format to apply to the cell.</param>
    private void WriteDateOnlyToCellWithUseSharedStrings(string value, Cell cell, ExcelCellFormats? excelcellformat)
    {
        if (DateOnly.TryParse(value, out DateOnly dateonlyvalue))
        {
            int index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart, SharedStringDictionary);

            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
        else
        {
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            cell.CellValue = new CellValue(value);
            if (excelcellformat is null)
                cell.StyleIndex = (uint)ExcelCellFormats.DateDefault;
            else
                cell.StyleIndex = (uint)excelcellformat;
        }
    }

    /// <summary>
    /// Writes a TimeOnly value to a cell.
    /// </summary>
    /// <param name="value">The TimeOnly value to write.</param>
    /// <param name="cell">The cell to write the value to.</param>
    /// <param name="excelcellformat">The format to apply to the cell.</param>
    private void WriteTimeOnlyToCellWithUseSharedStrings(string value, Cell cell, ExcelCellFormats? excelcellformat)
    {
        if (TimeOnly.TryParse(value, out TimeOnly timeonlyvalue))
        {
            int index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart, SharedStringDictionary);

            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
        else
        {
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            cell.CellValue = new CellValue(value);
            if (excelcellformat is null)
                cell.StyleIndex = (uint)ExcelCellFormats.TimeWithHoursMinutesSecondsDefault;
            else
                cell.StyleIndex = (uint)excelcellformat;
        }
    }

    /// <summary>
    /// Writes a DateTime value to a cell.
    /// </summary>
    /// <param name="value">The DateTime value to write.</param>
    /// <param name="Cell">The cell to write the value to.</param>
    /// <param name="ExcelCellFormat">The format to apply to the cell.</param>
    private void WriteDateTimeToCellWithUseSharedStrings(string value, Cell Cell, ExcelCellFormats? ExcelCellFormat)
    {
        if (DateTime.TryParse(value, out DateTime datetimevalue))
        {
            int index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart, SharedStringDictionary);

            Cell.CellValue = new CellValue(index.ToString());
            Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
        else
        {
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            Cell.CellValue = new CellValue(value);
            if (ExcelCellFormat is null)
                Cell.StyleIndex = (uint)ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault;
            else
                Cell.StyleIndex = (uint)ExcelCellFormat;
        }
    }

    /// <summary>
    /// Writes an integer value to a cell.
    /// </summary>
    /// <param name="value">The integer value to write.</param>
    /// <param name="cell">The cell to write the value to.</param>
    private void WriteIntToCell(string value, Cell cell)
    {
        cell.CellValue = new CellValue(value);
        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        cell.StyleIndex = (uint)ExcelCellFormats.NumberIntegerDefault;
    }

    /// <summary>
    /// Writes a decimal value to a cell.
    /// </summary>
    /// <param name="value">The decimal value to write.</param>
    /// <param name="cell">The cell to write the value to.</param>
    /// <param name="excelcellformat">The format to apply to the cell.</param>
    private void WriteDecimalToCell(string value, Cell cell, ExcelCellFormats? excelcellformat)
    {
        cell.CellValue = new CellValue(decimal.Parse(value, Configuration.CultureInfo));
        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        if (excelcellformat is null)
            cell.StyleIndex = (uint)ExcelCellFormats.NumberDecimalWithTwoDecimalsDefault;
        else
            cell.StyleIndex = (uint)excelcellformat;
    }

    /// <summary>
    /// Writes a double value to a cell.
    /// </summary>
    /// <param name="value">The double value to write.</param>
    /// <param name="cell">The cell to write the value to.</param>
    /// <param name="excelcellformat">The format to apply to the cell.</param>
    private void WriteDoubleToCell(string value, Cell cell, ExcelCellFormats? excelcellformat)
    {
        cell.CellValue = new CellValue(double.Parse(value, Configuration.CultureInfo));
        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        if (excelcellformat is null)
            cell.StyleIndex = (uint)ExcelCellFormats.ScientificWithTwoDecimalsDefault;
        else
            cell.StyleIndex = (uint)excelcellformat;
    }

    /// <summary>
    /// Writes a boolean value to a cell.
    /// </summary>
    /// <param name="value">The boolean value to write, represented as a string.</param>
    /// <param name="cell">The cell to write the value to.</param>
    private void WriteBoolToCell(string value, Cell cell)
    {
        cell.CellValue = new CellValue((bool.Parse(value) ? 1 : 0).ToString());
        cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
    }

    /// <summary>
    /// Adjusts the column widths to fit the content.
    /// </summary>
    private void AutoFitColumns()
    {
        if (ExcelCellMemberMapDetails.Count == 0)
            return;

        Columns columns = WorksheetPart.Worksheet.GetFirstChild<Columns>() ?? WorksheetPart.Worksheet.InsertAt(new Columns(), 0);

        for (int ColumnIndex = 0; ColumnIndex < ExcelColumnCount; ColumnIndex++)
        {
            double ColumnWidth = ExcelCellMemberMapDetails[ColumnIndex] switch
            {
                (nameof(DateOnly), null, <= 10) => 12,
                (nameof(DateOnly), ExcelCellFormats.DateDefault, <= 10) => 12,
                (nameof(DateOnly), ExcelCellFormats.DateExtended, <= 27) => 29,
                (nameof(DateOnly), ExcelCellFormats.DateWithDash, <= 10) => 12,
                (nameof(DateOnly), null, > 10) or (nameof(DateOnly), ExcelCellFormats.DateDefault, > 10) or (nameof(DateOnly), ExcelCellFormats.DateWithDash, > 10) or (nameof(DateOnly), ExcelCellFormats.DateExtended, > 27) => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 2,

                (nameof(TimeOnly), null, <= 8) => 10,
                (nameof(TimeOnly), ExcelCellFormats.TimeWithHoursMinutesSecondsDefault, <= 8) => 10,
                (nameof(TimeOnly), ExcelCellFormats.TimeWithHoursMinutes, <= 5) => 7,
                (nameof(TimeOnly), ExcelCellFormats.Time12HourWithHoursMinutesSeconds, <= 11) => 13,
                (nameof(TimeOnly), ExcelCellFormats.Time12HourWithHoursMinutes, <= 8) => 10,
                (nameof(TimeOnly), null, > 8) or (nameof(DateTime), ExcelCellFormats.TimeWithHoursMinutesSecondsDefault, > 8) or (nameof(DateTime), ExcelCellFormats.TimeWithHoursMinutes, > 5) or (nameof(DateTime), ExcelCellFormats.Time12HourWithHoursMinutesSeconds, > 11) or (nameof(DateTime), ExcelCellFormats.Time12HourWithHoursMinutes, > 8) => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 2,

                (nameof(DateTime), null, <= 19) => 21,
                (nameof(DateTime), ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault, <= 19) => 21,
                (nameof(DateTime), ExcelCellFormats.DateTimeWithHoursMinutes, <= 16) => 18,
                (nameof(DateTime), ExcelCellFormats.DateTime12HourWithHoursMinutesSeconds, <= 22) => 24,
                (nameof(DateTime), ExcelCellFormats.DateTime12HourWithHoursMinutes, <= 19) => 21,
                (nameof(DateTime), null, > 19) or (nameof(DateTime), ExcelCellFormats.DateTimeWithHoursMinutesSecondsDefault, > 19) or (nameof(DateTime), ExcelCellFormats.DateTimeWithHoursMinutes, > 16) or (nameof(DateTime), ExcelCellFormats.DateTime12HourWithHoursMinutesSeconds, > 22) or (nameof(DateTime), ExcelCellFormats.DateTime12HourWithHoursMinutes, > 19) => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 2,

                (nameof(Double), null, <= 8) => 10,
                (nameof(Double), ExcelCellFormats.ScientificWithTwoDecimalsDefault, <= 8) => 10,
                (nameof(Double), ExcelCellFormats.ScientificWithFourDecimals, <= 10) => 12,
                (nameof(Double), null, > 8) or (nameof(Double), ExcelCellFormats.ScientificWithTwoDecimalsDefault, > 8) or (nameof(Double), ExcelCellFormats.ScientificWithFourDecimals, > 10) => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 1,

                (_, null, < 10) => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 2,
                (_, _, >= 10 and <= 20) => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 4,
                (_, null, _) => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 1,

                _ => ExcelCellMemberMapDetails[ColumnIndex].CellLength + 2,
            };

            Column column = new Column() { Min = Convert.ToUInt32(ColumnIndex + 1), Max = Convert.ToUInt32(ColumnIndex + 1), Width = ColumnWidth, CustomWidth = true, BestFit = true };

            columns.Append(column);
        }

        WorksheetPart.Worksheet.Save();
    }

    private Type GetGenericType<T>(IEnumerable<T> enumerable)
    {
        return typeof(T);
    }

    #endregion
}