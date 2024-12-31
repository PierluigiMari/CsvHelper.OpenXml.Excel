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
    /// Cell format for Euro IT currency with two decimal places.
    /// </summary>
    CurrencyEuroITWithTwoDecimals = 9,
    /// <summary>
    /// Cell format for Euro IT currency with four decimal places.
    /// </summary>
    CurrencyEuroITWithFourDecimals = 10,
    /// <summary>
    /// Cell format for Dollar US currency with two decimal places.
    /// </summary>
    CurrencyDollarUSWithTwoDecimals = 11,
    /// <summary>
    /// Cell format for Dollar uS currency with four decimal places.
    /// </summary>
    CurrencyDollarUSWithFourDecimals = 12,
    /// <summary>
    /// Cell format for Pound GB currency with two decimal places.
    /// </summary>
    CurrencyPoundGBWithTwoDecimals = 13,
    /// <summary>
    /// Cell format for Pound GB currency with four decimal places.
    /// </summary>
    CurrencyPoundGBWithFourDecimals = 14,
    /// <summary>
    /// Cell format for Euro IT accounting with two decimal places.
    /// </summary>
    AccountingEuroITWithTwoDecimals = 15,
    /// <summary>
    /// Cell format for Euro IT accounting with four decimal places.
    /// </summary>
    AccountingEuroITWithFourDecimals = 16,
    /// <summary>
    /// Cell format for Dollar US accounting with two decimal places.
    /// </summary>
    AccountingDollarUSWithTwoDecimals = 17,
    /// <summary>
    /// Cell format for Dollar US accounting with four decimal places.
    /// </summary>
    AccountingDollarUSWithFourDecimals = 18,
    /// <summary>
    /// Cell format for Pound GB accounting with two decimal places.
    /// </summary>
    AccountingPoundGBWithTwoDecimals = 19,
    /// <summary>
    /// Cell format for Pound GB accounting with four decimal places.
    /// </summary>
    AccountingPoundGBWithFourDecimals = 20,
    /// <summary>
    /// Default cell format for dates.
    /// </summary>
    DateDefault = 21,
    /// <summary>
    /// Extended cell format for dates.
    /// </summary>
    DateExtended = 22,
    /// <summary>
    /// Cell format for dates with dashes.
    /// </summary>
    DateWithDash = 23,
    /// <summary>
    /// Default cell format for date and time with hours, minutes, and seconds.
    /// </summary>
    DateTimeWithHoursMinutesSecondsDefault = 24,
    /// <summary>
    /// Cell format for date and time with hours and minutes.
    /// </summary>
    DateTimeWithHoursMinutes = 25,
    /// <summary>
    /// Cell format for 12-hour date and time with hours, minutes, and seconds.
    /// </summary>
    DateTime12HourWithHoursMinutesSeconds = 26,
    /// <summary>
    /// Cell format for 12-hour date and time with hours and minutes.
    /// </summary>
    DateTime12HourWithHoursMinutes = 27,
    /// <summary>
    /// Default cell format for time with hours, minutes, and seconds.
    /// </summary>
    TimeWithHoursMinutesSecondsDefault = 28,
    /// <summary>
    /// Cell format for time with hours and minutes.
    /// </summary>
    TimeWithHoursMinutes = 29,
    /// <summary>
    /// Cell format for 12-hour time with hours, minutes, and seconds.
    /// </summary>
    Time12HourWithHoursMinutesSeconds = 30,
    /// <summary>
    /// Cell format for 12-hour time with hours and minutes.
    /// </summary>
    Time12HourWithHoursMinutes = 31,
    /// <summary>
    /// Cell format for percentages without decimal places.
    /// </summary>
    PercentageWithoutDecimals = 32,
    /// <summary>
    /// Cell format for percentages with two decimal places.
    /// </summary>
    PercentageWithTwoDecimals = 33,
    /// <summary>
    /// Default cell format for scientific notation with two decimal places.
    /// </summary>
    ScientificWithTwoDecimalsDefault = 34,
    /// <summary>
    /// Cell format for scientific notation with four decimal places.
    /// </summary>
    ScientificWithFourDecimals = 35,
    /// <summary>
    /// Cell format for special zip codes.
    /// </summary>
    SpecialZipCode = 36
}