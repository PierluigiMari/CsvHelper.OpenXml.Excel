namespace CsvHelper.OpenXml.Excel.TypeConversion;

public enum ExcelCellFormats : int
{
    Default = 0,
    DefaultBold = 1,
    DefaultBoldCentered = 2,
    NumberIntegerDefault = 3,
    NumberDecimalWithTwoDecimalsDefault = 4,
    NumberDecimalWithFourDecimals = 5,
    CurrencyGenericWithoutDecimals = 6,
    CurrencyGenericWithTwoDecimals = 7,
    CurrencyGenericWithFourDecimals = 8,
    CurrencyEuroWithTwoDecimals = 9,
    CurrencyEuroWithFourDecimals = 10,
    AccountingWithTwoDecimals = 11,
    AccountingWithFourDecimals = 12,
    DateDefault = 13,
    DateExtended = 14,
    DateWithDash = 15,
    DateTimeWithHoursMinutesSecondsDefault = 16,
    DateTimeWithHoursMinutes = 17,
    DateTime12HourWithHoursMinutesSeconds = 18,
    DateTime12HourWithHoursMinutes = 19,
    TimeWithHoursMinutesSecondsDefault = 20,
    TimeWithHoursMinutes = 21,
    Time12HourWithHoursMinutesSeconds = 22,
    Time12HourWithHoursMinutes = 23,
    PercentageWithoutDecimals = 24,
    PercentageWithTwoDecimals = 25,
    ScientificWithTwoDecimalsDefault = 26,
    ScientificWithFourDecimals = 27,
    SpecialZipCode = 28
}