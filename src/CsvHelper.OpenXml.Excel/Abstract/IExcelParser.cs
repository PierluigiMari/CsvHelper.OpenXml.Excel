namespace CsvHelper.OpenXml.Excel.Abstract;

using System;

public interface IExcelParser : IParser, IAsyncDisposable
{
    string[] GetRecord();
}