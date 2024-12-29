namespace CsvHelper.OpenXml.Excel.Abstractions;

using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;

/// <summary>
/// Defines the contract for implementing a ExcelWriter for writing records to an Excel file.
/// </summary>
public interface IExcelWriter : IWriter
{
    /// <summary>
    /// Writes a single record to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    /// <param name="record">The record to write.</param>
    /// <param name="sheetname">The name of the sheet to write to.</param>
    void WriteRecord<T>(T? record, string? sheetname);

    /// <summary>
    /// Writes multiple records to the specified sheet.
    /// </summary>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to. If null, the default sheet is used.</param>
    void WriteRecords(IEnumerable records, string? sheetname = null);
    /// <summary>
    /// Asynchronously writes multiple records to the specified sheet.
    /// </summary>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to. If null, the default sheet is used.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    Task WriteRecordsAsync(IEnumerable records, string? sheetname = null, CancellationToken cancellationToken = default);

    /// <summary>
    /// Writes multiple records of a specific type to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the records.</typeparam>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to.</param>
    void WriteRecords<T>(IEnumerable<T> records, string? sheetname);
    /// <summary>
    /// Asynchronously writes multiple records of a specific type to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the records.</typeparam>
    /// <param name="records">The records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    Task WriteRecordsAsync<T>(IEnumerable<T> records, string? sheetname, CancellationToken cancellationToken = default);

    /// <summary>
    /// Asynchronously writes multiple records from an asynchronous enumerable to the specified sheet.
    /// </summary>
    /// <typeparam name="T">The type of the records.</typeparam>
    /// <param name="records">The asynchronous enumerable of records to write.</param>
    /// <param name="sheetname">The name of the sheet to write to.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    Task WriteRecordsAsync<T>(IAsyncEnumerable<T> records, string? sheetname, CancellationToken cancellationToken = default);
}