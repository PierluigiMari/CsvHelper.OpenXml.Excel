namespace CsvHelper.OpenXml.Excel.Tests;

using System;

internal class Record
{
    public required int Id { get; set; }
    public required string Number { get; set; }
    public string? Description { get; set; }
    public DateTime? Date { get; set; }
    public DateTime? AnotherDate { get; set; }
    public DateTime? YetAnotherDate { get; set; }
    public required decimal Amount { get; set; }
    public required DateOnly CreationDate { get; set; }
    public required TimeOnly CreationTime { get; set; }
    public DateTime? LastModifiedDate { get; set; }
}