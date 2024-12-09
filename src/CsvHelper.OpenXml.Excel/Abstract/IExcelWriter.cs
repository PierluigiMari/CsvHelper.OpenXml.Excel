namespace CsvHelper.OpenXml.Excel.Abstract;

using System.Collections.Generic;
using System.Threading.Tasks;

public interface IExcelWriter : IWriter
{
    void WriteRecord<T>(T? record, string? sheetname);

    void WriteRecords<T>(IEnumerable<T> records, string? sheetname);
    Task WriteRecordsAsync<T>(IEnumerable<T> records, string? sheetname, CancellationToken cancellationToken = default);

    Task WriteRecordsAsync<T>(IAsyncEnumerable<T> records, string? sheetname, CancellationToken cancellationToken = default);
}