namespace CsvHelper.OpenXml.Excel.Abstractions;

using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;

public interface IExcelWriter : IWriter
{
    void WriteRecord<T>(T? record, string? sheetname);

    void WriteRecords(IEnumerable records, string? sheetname = null);
    Task WriteRecordsAsync(IEnumerable records, string? sheetname = null, CancellationToken cancellationToken = default);

    void WriteRecords<T>(IEnumerable<T> records, string? sheetname);
    Task WriteRecordsAsync<T>(IEnumerable<T> records, string? sheetname, CancellationToken cancellationToken = default);

    Task WriteRecordsAsync<T>(IAsyncEnumerable<T> records, string? sheetname, CancellationToken cancellationToken = default);
}