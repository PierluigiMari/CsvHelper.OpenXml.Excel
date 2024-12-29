namespace CsvHelper.OpenXml.Excel.Abstractions;

using System;

/// <summary>
/// Defines the contract for implementing a ExcelParser for parsing Excel files.
/// </summary>
public interface IExcelParser : IParser, IAsyncDisposable
{
    /// <summary>
    /// Gets the total number of rows in the Excel file.
    /// </summary>
    int RowCount { get; }

    /// <summary>
    /// Retrieves the record as an array of strings.
    /// </summary>
    /// <returns>An array of strings representing the record.</returns>
    string[] GetRecord();
}