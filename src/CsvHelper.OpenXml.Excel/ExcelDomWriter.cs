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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

public sealed class ExcelDomWriter : CsvWriter, IExcelWriter
{
    #region Fields

    private readonly OpenXmlHelper OpenXmlHelper = new OpenXmlHelper();

    private readonly InjectionOptions InjectionOptions;

    private readonly SpreadsheetDocument SpreadsheetDocument;

    private readonly WorkbookPart WorkbookPart;

    private readonly SharedStringTablePart SharedStringPart;

    private WorksheetPart WorksheetPart = null!;

    private SheetData SheetData = null!;

    private int ExcelRowIndex = 1;
    private int ExcelColumnIndex = 0;
    private int ExcelColumnCount = 0;

    private Row WritingRow = null!;

    private Type? WritingFieldType = null!;

    private readonly Dictionary<int, (string FieldTypeName, ExcelCellFormats? ExcelCellFormat, double CellLength)> ExcelCellMemberMapDetails = new Dictionary<int, (string FieldTypeName, ExcelCellFormats? ExcelCellFormat, double CellLength)>();

    private string? LastSheetName;

    #endregion

    #region Constructors

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
    /// </summary>
    /// <param name="stream">The stream.</param>
    /// <param name="culture">The culture.</param>
    public ExcelDomWriter(Stream stream, CultureInfo? culture) : this(stream, culture is null ? new CsvConfiguration(CultureInfo.InvariantCulture) : new CsvConfiguration(culture)) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
    /// </summary>
    /// <param name="stream">The stream.</param>
    /// <param name="configuration">The configuration.</param>
    public ExcelDomWriter(Stream stream, CsvConfiguration? configuration = null) : base(TextWriter.Null, configuration is null ? new CsvConfiguration(CultureInfo.InvariantCulture) : configuration)
    {
        base.Configuration.Validate();

        if (stream.Length > 0)
        {
            SpreadsheetDocument = SpreadsheetDocument.Open(stream, true);
        }
        else
            SpreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart = SpreadsheetDocument.WorkbookPart ?? SpreadsheetDocument.AddWorkbookPart();

        WorkbookPart.Workbook ??= new Workbook();

        OpenXmlHelper.CreateWorksheetStyle(SpreadsheetDocument);

        SharedStringPart = OpenXmlHelper.GetSharedStringTablePart(WorkbookPart);

        InjectionOptions = base.Configuration.InjectionOptions;
    }

    #endregion

    #region Overriding CsvWriter Properties and Methods

    public override int Row => ExcelRowIndex;
    public override int Index => ExcelColumnIndex;


    public override void WriteField<T>(T? field, ITypeConverter converter) where T : default
    {
        WritingFieldType = typeof(T);

        base.WriteField(field, converter);
    }

    public override void WriteConvertedField(string? field, Type fieldType)
    {
        WritingFieldType = fieldType;

        base.WriteConvertedField(field, fieldType);
    }

    public override void WriteField(string? field, bool shouldQuote)
    {
        if (InjectionOptions == InjectionOptions.Strip)
        {
            field = SanitizeForInjection(field);
        }

        WriteToCell(field);

        ExcelColumnIndex++;
    }


    public override void NextRecord()
    {
        ExcelColumnCount = ExcelColumnIndex;
        ExcelColumnIndex = 0;
        ExcelRowIndex++;

        WritingRow = new Row();
        SheetData.Append(WritingRow);
    }

    public override Task NextRecordAsync()
    {
        NextRecord();

        return Task.CompletedTask;
    }

    public override void Flush()
    {
        WorkbookPart.Workbook.Save();
    }

    public override Task FlushAsync()
    {
        WorkbookPart.Workbook.Save();

        return Task.CompletedTask;
    }

    #region IDisposable and IAsyncDisposable override Methods of CsvWriter

    private bool Disposed;

    protected override void Dispose(bool disposing)
    {
        if (Disposed)
            return;

        WorkbookPart.Workbook.Save();

        if (disposing)
        {
            SpreadsheetDocument.Dispose();
        }

        // TODO: free unmanaged resources (unmanaged objects) and override finalizer
        // TODO: set large fields to null
        Disposed = true;
    }

    protected override ValueTask DisposeAsync(bool disposing)
    {
        if (Disposed)
            return ValueTask.CompletedTask;

        WorkbookPart.Workbook.Save();

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

    public void WriteRecord<T>(T? record, string? sheetname = null)
    {
        if (ExcelRowIndex == 1)
        {
            InitializeWritingNewWorksheet<T>(sheetname);

            WriteHeader<T>();

            NextRecord();
        }

        base.WriteRecord(record);
    }


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

        AutoFitColumns();
    }

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

        AutoFitColumns();
    }


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

        AutoFitColumns();
    }

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

        AutoFitColumns();
    }


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

        AutoFitColumns();
    }

    #endregion

    #region Private Methods

    private void InitializeWritingNewWorksheet(Type type, string? sheetname)
    {
        WorksheetPart = OpenXmlHelper.InsertWorksheet(WorkbookPart, string.IsNullOrEmpty(sheetname) ? null : sheetname);

        SheetData = WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;

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

    private void InitializeWritingNewWorksheet<T>(string? sheetname)
    {
        WorksheetPart = OpenXmlHelper.InsertWorksheet(WorkbookPart, string.IsNullOrEmpty(sheetname) ? null : sheetname);

        //WorksheetPart.Worksheet.InsertAt(new Columns(), 0);

        SheetData = WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;

        ClassMap? ClassMap = Context.Maps[typeof(T)];
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

    private void WriteToCell(string? value)
    {
        int length = value?.Length ?? 0;

        if (value is null || length == 0)
        {
            WritingFieldType = null;
            return;
        }

        if (ExcelCellMemberMapDetails[ExcelColumnIndex].CellLength < value.Length)
            ExcelCellMemberMapDetails[ExcelColumnIndex] = (ExcelCellMemberMapDetails[ExcelColumnIndex].FieldTypeName, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat, value.Length);

        Cell Cell = new Cell() { CellReference = $"{OpenXmlHelper.GetColumnLetters(ExcelColumnIndex)}{ExcelRowIndex}" };
        WritingRow.Append(Cell);

        if (WritingFieldType is not null)
        {
            Action WriteSpecificTypeInCell = WritingFieldType.Name switch
            {
                nameof(String) or nameof(Guid) => () => WriteStringOrGuidToCell(value, Cell, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat),
                nameof(DateOnly) => () => WriteDateOnlyToCell(value, Cell, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat),
                nameof(TimeOnly) => () => WriteTimeOnlyToCell(value, Cell, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat),
                nameof(DateTime) => () => WriteDateTimeToCell(value, Cell, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat),
                nameof(Int32) => () => WriteIntTocell(value, Cell),
                nameof(Decimal) => () => WriteDecimalToCell(value, Cell, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat),
                nameof(Double) => () => WriteDoubleToCell(value, Cell, ExcelCellMemberMapDetails[ExcelColumnIndex].ExcelCellFormat),
                nameof(Boolean) => () => WriteBoolToCell(value, Cell),
                _ => throw new NotImplementedException($"Writing of the specific type {WritingFieldType.Name} not yet implemented!")
            };

            WriteSpecificTypeInCell();
        }
        else
        {
            WriteStringOrGuidToCell(value, Cell, null, true);
        }

        WritingFieldType = null;

        void WriteStringOrGuidToCell(string value, Cell Cell, ExcelCellFormats? ExcelCellFormat, bool isheadercell = false)
        {
            if (int.TryParse(value, out _))
            {
                Cell.CellValue = new CellValue(value);

                if (ExcelCellFormat is not null)
                    Cell.StyleIndex = (uint)ExcelCellFormat;

            }
            else
            {
                int index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart);

                Cell.CellValue = new CellValue(index.ToString());
                Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                if (isheadercell)
                    Cell.StyleIndex = (uint)ExcelCellFormats.DefaultBoldCentered;
            }
        }

        void WriteDateOnlyToCell(string value, Cell Cell, ExcelCellFormats? ExcelCellFormat)
        {
            if (DateOnly.TryParse(value, out DateOnly dateonlyvalue))
            {
                int index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart);

                Cell.CellValue = new CellValue(index.ToString());
                Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }
            else
            {
                Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                Cell.CellValue = new CellValue(value);
                if (ExcelCellFormat is null)
                    Cell.StyleIndex = (uint)ExcelCellFormats.DateDefault;
                else
                    Cell.StyleIndex = (uint)ExcelCellFormat;
            }
        }

        void WriteTimeOnlyToCell(string value, Cell Cell, ExcelCellFormats? ExcelCellFormat)
        {
            if (TimeOnly.TryParse(value, out TimeOnly timeonlyvalue))
            {
                int index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart);

                Cell.CellValue = new CellValue(index.ToString());
                Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }
            else
            {
                Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                Cell.CellValue = new CellValue(value);
                if (ExcelCellFormat is null)
                    Cell.StyleIndex = (uint)ExcelCellFormats.TimeWithHoursMinutesSecondsDefault;
                else
                    Cell.StyleIndex = (uint)ExcelCellFormat;
            }
        }

        void WriteDateTimeToCell(string value, Cell Cell, ExcelCellFormats? ExcelCellFormat)
        {
            if (DateTime.TryParse(value, out DateTime datetimevalue))
            {
                int index = OpenXmlHelper.InsertSharedStringItem(value, SharedStringPart);

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

        void WriteIntTocell(string value, Cell Cell)
        {
            Cell.CellValue = new CellValue(value);
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            Cell.StyleIndex = (uint)ExcelCellFormats.NumberIntegerDefault;
        }

        void WriteDecimalToCell(string value, Cell Cell, ExcelCellFormats? ExcelCellFormat)
        {
            Cell.CellValue = new CellValue(decimal.Parse(value, Configuration.CultureInfo));
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            if (ExcelCellFormat is null)
                Cell.StyleIndex = (uint)ExcelCellFormats.NumberDecimalWithTwoDecimalsDefault;
            else
                Cell.StyleIndex = (uint)ExcelCellFormat;
        }

        void WriteDoubleToCell(string value, Cell Cell, ExcelCellFormats? ExcelCellFormat)
        {
            Cell.CellValue = new CellValue(double.Parse(value, Configuration.CultureInfo));
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            if (ExcelCellFormat is null)
                Cell.StyleIndex = (uint)ExcelCellFormats.ScientificWithTwoDecimalsDefault;
            else
                Cell.StyleIndex = (uint)ExcelCellFormat;
        }

        void WriteBoolToCell(string value, Cell Cell)
        {
            Cell.CellValue = new CellValue((bool.Parse(value) ? 1 : 0).ToString());
            Cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
    }

    private void AutoFitColumns()
    {
        WorksheetPart.Worksheet.InsertAt(new Columns(), 0);

        Columns columns = WorksheetPart.Worksheet.GetFirstChild<Columns>()!;

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

            Column column = new Column() { Min = Convert.ToUInt32(ColumnIndex + 1), Max = Convert.ToUInt32(ColumnIndex + 1), Width = ColumnWidth, CustomWidth = BooleanValue.FromBoolean(true), BestFit = BooleanValue.FromBoolean(true) };

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