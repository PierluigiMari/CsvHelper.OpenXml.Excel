namespace CsvHelper.OpenXml.Excel.TypeConversion;

/// <summary>
/// Enumeration that specifies the various cell formats that can be applied to Excel cells.
/// </summary>
public enum ExcelCellFormats : int
{
    /// <summary>
    /// Default cell format.
    /// </summary>
    Default = 0,
    /// <summary>
    /// Default cell format with bold text.
    /// </summary>
    DefaultBold = 1,
    /// <summary>
    /// Default cell format with bold and centered text.
    /// </summary>
    DefaultBoldCentered = 2,
    /// <summary>
    /// Cell format for integer numbers.
    /// </summary>
    NumberIntegerDefault = 3,
    /// <summary>
    /// Cell format for decimal numbers with two decimal places.
    /// </summary>
    NumberDecimalWithTwoDecimalsDefault = 4,
    /// <summary>
    /// Cell format for decimal numbers with four decimal places.
    /// </summary>
    NumberDecimalWithFourDecimals = 5,
    /// <summary>
    /// Cell format for generic currency without decimal places.
    /// </summary>
    CurrencyGenericWithoutDecimals = 6,
    /// <summary>
    /// Cell format for generic currency with two decimal places.
    /// </summary>
    CurrencyGenericWithTwoDecimals = 7,
    /// <summary>
    /// Cell format for generic currency with four decimal places.
    /// </summary>
    CurrencyGenericWithFourDecimals = 8,
    /// <summary>
    /// Cell format for Euro currency with two decimal places.
    /// </summary>
    CurrencyEuroWithTwoDecimals = 9,
    /// <summary>
    /// Cell format for Euro currency with four decimal places.
    /// </summary>
    CurrencyEuroWithFourDecimals = 10,
    /// <summary>
    /// Cell format for accounting with two decimal places.
    /// </summary>
    AccountingWithTwoDecimals = 11,
    /// <summary>
    /// Cell format for accounting with four decimal places.
    /// </summary>
    AccountingWithFourDecimals = 12,
    /// <summary>
    /// Default cell format for dates.
    /// </summary>
    DateDefault = 13,
    /// <summary>
    /// Extended cell format for dates.
    /// </summary>
    DateExtended = 14,
    /// <summary>
    /// Cell format for dates with dashes.
    /// </summary>
    DateWithDash = 15,
    /// <summary>
    /// Default cell format for date and time with hours, minutes, and seconds.
    /// </summary>
    DateTimeWithHoursMinutesSecondsDefault = 16,
    /// <summary>
    /// Cell format for date and time with hours and minutes.
    /// </summary>
    DateTimeWithHoursMinutes = 17,
    /// <summary>
    /// Cell format for 12-hour date and time with hours, minutes, and seconds.
    /// </summary>
    DateTime12HourWithHoursMinutesSeconds = 18,
    /// <summary>
    /// Cell format for 12-hour date and time with hours and minutes.
    /// </summary>
    DateTime12HourWithHoursMinutes = 19,
    /// <summary>
    /// Default cell format for time with hours, minutes, and seconds.
    /// </summary>
    TimeWithHoursMinutesSecondsDefault = 20,
    /// <summary>
    /// Cell format for time with hours and minutes.
    /// </summary>
    TimeWithHoursMinutes = 21,
    /// <summary>
    /// Cell format for 12-hour time with hours, minutes, and seconds.
    /// </summary>
    Time12HourWithHoursMinutesSeconds = 22,
    /// <summary>
    /// Cell format for 12-hour time with hours and minutes.
    /// </summary>
    Time12HourWithHoursMinutes = 23,
    /// <summary>
    /// Cell format for percentages without decimal places.
    /// </summary>
    PercentageWithoutDecimals = 24,
    /// <summary>
    /// Cell format for percentages with two decimal places.
    /// </summary>
    PercentageWithTwoDecimals = 25,
    /// <summary>
    /// Default cell format for scientific notation with two decimal places.
    /// </summary>
    ScientificWithTwoDecimalsDefault = 26,
    /// <summary>
    /// Cell format for scientific notation with four decimal places.
    /// </summary>
    ScientificWithFourDecimals = 27,
    /// <summary>
    /// Cell format for special zip codes.
    /// </summary>
    SpecialZipCode = 28
}