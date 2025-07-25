﻿namespace CsvHelper.OpenXml.Excel.Tests.ConsoleApp;

using System;
using System.Text.Json;

public class GenericPerson
{
    public Guid Id { get; set; }
    public string FirstName { get; set; } = null!;
    public string Lastname { get; set; } = null!;
    public int NumberInteger { get; set; }
    public decimal NumberDecimalWithTwoDecimals { get; set; }
    public decimal NumberDecimalWithFourDecimals { get; set; }
    public string PhoneNo { get; set; } = null!;
    //public (string Text, string EmailAddress) EmailId { get; set; }
    public string EmailId { get; set; } = null!;
    public string? Address { get; set; }
    public string ZipCode { get; set; } = null!;
    public DateOnly DateDefault { get; set; }
    public DateOnly DateExtendedDefault { get; set; }
    public bool Adult { get; set; }
    public DateOnly DateWithDash { get; set; }
    public decimal GenericCurrencyWithoutDecimals { get; set; }
    public decimal GenericCurrencyWithTwoDecimals { get; set; }
    public decimal GenericCurrencyWithFourDecimals { get; set; }
    public decimal EuroCurrencyWithTwoDecimals { get; set; }
    public decimal EuroCurrencyWithFourDecimals { get; set; }
    public decimal AccountingWithTwoDecimals { get; set; }
    public decimal AccountingWithFourDecimals { get; set; }
    public decimal Percentage { get; set; }
    public decimal PercentageWithTwoDecimal { get; set; }
    public TimeOnly TimeDefault { get; set; }
    public TimeOnly? TimeHoursMinutes { get; set; }
    public TimeOnly Time12HoursMinutesSeconds { get; set; }
    public TimeOnly Time12HoursMinutes { get; set; }
    public DateTime DateTimeDefault { get; set; }
    public DateTime? DateTimeWithHoursMinutes { get; set; }
    public DateTime DateTime12HourWithHoursMinutesSeconds { get; set; }
    public DateTime DateTime12HourWithHoursMinutes { get; set; }
    public DateTimeOffset DateTimeOffsetUnspecified { get; set; }
    public DateTimeOffset DateTimeOffsetUnspecifiedAsText { get; set; }
    public DateTimeOffset DateTimeOffsetUnspecifiedFromDateTimeAsText { get; set; }
    public DateTimeOffset DateTimeOffsetUtc { get; set; }
    public DateTimeOffset DateTimeOffsetUtcAsText { get; set; }
    public DateTimeOffset DateTimeOffsetLocal { get; set; }
    public DateTimeOffset DateTimeOffsetLocalAsText { get; set; }
    public double ScientificDefault { get; set; }
    public double? ScientificFourDecimal { get; set; }
    public string FileName { get; set; } = null!;

    //public (string, Uri) LinkText { get; set; }// = null!;
    //public Uri Link { get; set; } = null!;
    //public (string Text, string CellReference) LinkToCell { get; set; }
    //public string LinkToCell { get; set; } = null!;

    public override string ToString()
    {
        return JsonSerializer.Serialize(this, new JsonSerializerOptions { IncludeFields = true, WriteIndented = true });
    }
}