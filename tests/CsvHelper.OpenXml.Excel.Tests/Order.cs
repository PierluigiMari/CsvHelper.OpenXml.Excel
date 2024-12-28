namespace CsvHelper.OpenXml.Excel.Tests;

using System;

internal class Order
{
    public required int OrderId { get; set; }
    public required string OrderNumber { get; set; }
    public required DateOnly OrderDate { get; set; }
    public required TimeOnly OrderTime { get; set; }
    public required decimal OrderAmount { get; set; }
    public required string CustomerName { get; set; }
    public required string CustomerAddress { get; set; }
    public required string CustomerZip { get; set; }
    public required string CustomerCity { get; set; }
    public DateTime? ShippedDate { get; set; }
}