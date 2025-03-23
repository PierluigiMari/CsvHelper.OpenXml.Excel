namespace CsvHelper.OpenXml.Excel.Tests.ConsoleApp;

using System;
using System.Text.Json;

public class GenericOrder
{
    public int OrderId { get; set; }
    public string OrderNumber { get; set; } = null!;
    public DateOnly OrderDate { get; set; }
    public TimeOnly OrderTime { get; set; }
    public decimal OrderAmount { get; set; }
    public string CustomerName { get; set; } = null!;
    public string CustomerAddress { get; set; } = null!;
    public string CustomerZip { get; set; } = null!;
    public string CustomerCity { get; set; } = null!;
    public DateTime? ShippedDate { get; set; }

    public override string ToString()
    {
        return JsonSerializer.Serialize(this);
    }
}