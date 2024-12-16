namespace CsvHelper.OpenXml.Excel.Abstractions;

using System;

public interface IExcelParser : IParser, IAsyncDisposable
{
    int RowCount { get; }

    string[] GetRecord();
}