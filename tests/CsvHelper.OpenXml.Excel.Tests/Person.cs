namespace CsvHelper.OpenXml.Excel.Tests;

using System;

internal class Person
{
    public required string Name { get; set; }
    public required string Surname { get; set; }
    public string? NickName { get; set; }
    public required DateOnly BirthDate { get; set; }
    public required int Age { get; set; }
    public required DateTimeOffset WeddingDate { get; set; }
    public required string Address { get; set; }
    public required string Zip { get; set; }
    public required string City { get; set; }
    public required DateOnly CreationDate { get; set; }
    public required TimeOnly CreationTime { get; set; }
    public DateTime? LastModifiedDate { get; set; }
}