﻿using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("CsvHelper.OpenXml.Excel.Tests")]

namespace CsvHelper.OpenXml.Excel;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

/// <summary>
/// The OpenXmlHelper class is a utility class to assist with various operations related to OpenXML spreadsheets.
/// </summary>
internal partial class OpenXmlHelper
{
    /// <summary>
    /// Creates and applies a new stylesheet to the specified spreadsheet document.
    /// </summary>
    /// <param name="spreadsheetdocument">The spreadsheet document to which the stylesheet will be applied.</param>
    internal void CreateWorksheetStyle(SpreadsheetDocument spreadsheetdocument)
    {
        spreadsheetdocument.WorkbookPart!.DeletePart(spreadsheetdocument.WorkbookPart!.WorkbookStylesPart!);

        WorkbookStylesPart NewWorkbookStylesPartCreated = spreadsheetdocument.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
        Stylesheet WorkbookStyleSheet = new Stylesheet();

        NumberingFormats StylesheetNumberingFormats = new NumberingFormats();
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 164, FormatCode = "0.0000" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 165, FormatCode = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 166, FormatCode = "dd\\-mm\\-yyyy" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 167, FormatCode = "#,##0\\ _€" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 168, FormatCode = "#,##0.00\\ _€" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 169, FormatCode = "#,##0.0000\\ _€" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 170, FormatCode = "#,##0.00\\ [$€-410]" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 171, FormatCode = "#,##0.0000\\ [$€-410]" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 172, FormatCode = "[$$-409]#,##0.00" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 173, FormatCode = "[$$-409]#,##0.0000" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 174, FormatCode = "[$£-809]#,##0.00" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 175, FormatCode = "[$£-809]#,##0.0000" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 176, FormatCode = "_-* #,##0.00\\ [$€-410]_-;\\-* #,##0.00\\ [$€-410]_-;_-* \"-\"??\\ [$€-410]_-;_-@_-" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 177, FormatCode = "_-* #,##0.0000\\ [$€-410]_-;\\-* #,##0.0000\\ [$€-410]_-;_-* \"-\"????\\ [$€-410]_-;_-@_-" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 178, FormatCode = "_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* \"-\"??_ ;_-@_ " });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 179, FormatCode = "_-[$$-409]* #,##0.0000_ ;_-[$$-409]* \\-#,##0.0000\\ ;_-[$$-409]* \"-\"????_ ;_-@_ " });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 180, FormatCode = "_-[$£-809]* #,##0.00_-;\\-[$£-809]* #,##0.00_-;_-[$£-809]* \"-\"??_-;_-@_-" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 181, FormatCode = "_-[$£-809]* #,##0.0000_-;\\-[$£-809]* #,##0.0000_-;_-[$£-809]* \"-\"????_-;_-@_-" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 182, FormatCode = "[$-F400]h:mm:ss\\ AM/PM" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 183, FormatCode = "h:mm;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 184, FormatCode = "[$-409]h:mm:ss\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 185, FormatCode = "[$-409]h:mm\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 186, FormatCode = "dd/mm/yyyy\\ hh:mm:ss" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 187, FormatCode = "dd/mm/yyyy\\ h:mm:ss\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 188, FormatCode = "dd/mm/yyyy\\ h:mm\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 189, FormatCode = "0.0000E+00" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 190, FormatCode = "00000" });

        Font FontDefault = new Font(new FontName { Val = "Calibri" }, new FontSize { Val = 11 }); // Default font
        Font FontBold = new Font(new Bold()); // Bold font
        Font FontHyperlink = new Font(new Color { Theme = 10 }, new Underline { Val = UnderlineValues.Single }, new FontName { Val = "Calibri" }, new FontSize { Val = 11 }); // Hyperlink font
        //Font FontHyperlink = new Font(new Color { Theme = 10 }, new FontScheme { Val = FontSchemeValues.Minor }, new FontName { Val = "Calibri" }, new FontSize { Val = 11 }); // Hyperlink font

        Fonts Fonts = new Fonts(FontDefault, FontBold, FontHyperlink);

        Fill FillDefault = new Fill(new PatternFill { PatternType = PatternValues.None }); // Default fill
        Fills Fills = new Fills(FillDefault);

        Border BorderDefault = new Border(); // Default border
        Borders Borders = new Borders(BorderDefault);


        CellStyleFormats CellStyleFormats = new CellStyleFormats();
        CellStyleFormats.Append(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }); // Default cell format
        CellStyleFormats.Append(new CellFormat { NumberFormatId = 0, FontId = 2, FillId = 0, BorderId = 0, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false });


        CellStyles CellStyles = new CellStyles();
        CellStyles.Append(new CellStyle { Name = "Hyperlink", FormatId = 1, BuiltinId = 8 });
        CellStyles.Append(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 });

        // CellFormats
        CellFormats CellFormats = new CellFormats();

        CellFormat CellFormatDefault = new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }; // Default style : Mandatory | Style ID =0
        CellFormat CellFormatDefaultBold = new CellFormat { FontId = 1 };  // Style with Bold text ; Style ID = 1
        CellFormat CellFormatDefaultBoldCentered = new CellFormat { BorderId = 0, FillId = 0, FontId = 1, ApplyFont = true, ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center } }; //Style with Bold text with horizontal aligment centered
        CellFormat NumberIntegerFormat = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 1, FormatId = 0, ApplyNumberFormat = true };
        CellFormat NumberDecimalFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 2, FormatId = 0, ApplyNumberFormat = true };
        CellFormat NumberDecimalFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 164, FormatId = 0, ApplyNumberFormat = true };
        CellFormat CurrencyGenericFormatWithoutDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 167, FormatId = 0, ApplyNumberFormat = true }; // format like "#,##0"
        CellFormat CurrencyGenericFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 168, FormatId = 0, ApplyNumberFormat = true }; // format like "#,##0.00"
        CellFormat CurrencyGenericFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 169, FormatId = 0, ApplyNumberFormat = true };
        CellFormat CurrencyEuroITFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 170, FormatId = 0, ApplyNumberFormat = true }; // format like "#,##0.00 €"
        CellFormat CurrencyEuroITFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 171, FormatId = 0, ApplyNumberFormat = true };
        CellFormat CurrencyDollarUSFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 172, FormatId = 0, ApplyNumberFormat = true }; // format like "$#,##0.00"
        CellFormat CurrencyDollarUSFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 173, FormatId = 0, ApplyNumberFormat = true };
        CellFormat CurrencyPoundGBFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 174, FormatId = 0, ApplyNumberFormat = true }; // format like "£#,##0.00"
        CellFormat CurrencyPoundGBFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 175, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingEuroITFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 176, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingEuroITFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 177, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingDollarUSFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 178, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingDollarUSFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 179, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingPoundGBFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 180, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingPoundGBFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 181, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateFormatDefault = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 14, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateFormatExtended = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 165, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateFormatWithDash = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 166, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTimeFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 186, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTimeFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 22, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTime12HourFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 187, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTime12HourFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 188, FormatId = 0, ApplyNumberFormat = true };
        CellFormat TimeFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 182, FormatId = 0, ApplyNumberFormat = true };
        CellFormat TimeFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 183, FormatId = 0, ApplyNumberFormat = true };
        CellFormat Time12HourFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 184, FormatId = 0, ApplyNumberFormat = true };
        CellFormat Time12HourFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 185, FormatId = 0, ApplyNumberFormat = true };
        CellFormat PercentageFormatWithoutDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 9, FormatId = 0, ApplyNumberFormat = true };
        CellFormat PercentageFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 10, FormatId = 0, ApplyNumberFormat = true };
        CellFormat ScientificFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 11, FormatId = 0, ApplyNumberFormat = true };
        CellFormat ScientificFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 189, FormatId = 0, ApplyNumberFormat = true };
        CellFormat TextFormat = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 49, FormatId = 0, ApplyNumberFormat = true };
        CellFormat SpecialZipCodeFormat = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 190, FormatId = 0, ApplyNumberFormat = true };
        CellFormat HyperlinkFormat = new CellFormat { NumberFormatId = 0, FontId = 2, FillId = 0, BorderId = 0, FormatId = 1, ApplyFont = true };

        CellFormats.Append(CellFormatDefault);
        CellFormats.Append(CellFormatDefaultBold);
        CellFormats.Append(CellFormatDefaultBoldCentered);
        CellFormats.Append(NumberIntegerFormat);
        CellFormats.Append(NumberDecimalFormatWithTwoDecimals);
        CellFormats.Append(NumberDecimalFormatWithFourDecimals);
        CellFormats.Append(CurrencyGenericFormatWithoutDecimals);
        CellFormats.Append(CurrencyGenericFormatWithTwoDecimals);
        CellFormats.Append(CurrencyGenericFormatWithFourDecimals);
        CellFormats.Append(CurrencyEuroITFormatWithTwoDecimals);
        CellFormats.Append(CurrencyEuroITFormatWithFourDecimals);
        CellFormats.Append(CurrencyDollarUSFormatWithTwoDecimals);
        CellFormats.Append(CurrencyDollarUSFormatWithFourDecimals);
        CellFormats.Append(CurrencyPoundGBFormatWithTwoDecimals);
        CellFormats.Append(CurrencyPoundGBFormatWithFourDecimals);
        CellFormats.Append(AccountingEuroITFormatWithTwoDecimals);
        CellFormats.Append(AccountingEuroITFormatWithFourDecimals);
        CellFormats.Append(AccountingDollarUSFormatWithTwoDecimals);
        CellFormats.Append(AccountingDollarUSFormatWithFourDecimals);
        CellFormats.Append(AccountingPoundGBFormatWithTwoDecimals);
        CellFormats.Append(AccountingPoundGBFormatWithFourDecimals);
        CellFormats.Append(DateFormatDefault);
        CellFormats.Append(DateFormatExtended);
        CellFormats.Append(DateFormatWithDash);
        CellFormats.Append(DateTimeFormatWithHoursMinutesSeconds);
        CellFormats.Append(DateTimeFormatWithHoursMinutes);
        CellFormats.Append(DateTime12HourFormatWithHoursMinutesSeconds);
        CellFormats.Append(DateTime12HourFormatWithHoursMinutes);
        CellFormats.Append(TimeFormatWithHoursMinutesSeconds);
        CellFormats.Append(TimeFormatWithHoursMinutes);
        CellFormats.Append(Time12HourFormatWithHoursMinutesSeconds);
        CellFormats.Append(Time12HourFormatWithHoursMinutes);
        CellFormats.Append(PercentageFormatWithoutDecimals);
        CellFormats.Append(PercentageFormatWithTwoDecimals);
        CellFormats.Append(ScientificFormatWithTwoDecimals);
        CellFormats.Append(ScientificFormatWithFourDecimals);
        CellFormats.Append(TextFormat);
        CellFormats.Append(SpecialZipCodeFormat);
        CellFormats.Append(HyperlinkFormat);

        // Append everything to stylesheet  - Preserve the ORDER!
        WorkbookStyleSheet.Append(StylesheetNumberingFormats);
        WorkbookStyleSheet.Append(Fonts);
        WorkbookStyleSheet.Append(Fills);
        WorkbookStyleSheet.Append(Borders);
        WorkbookStyleSheet.Append(CellStyleFormats);
        WorkbookStyleSheet.Append(CellFormats);
        WorkbookStyleSheet.Append(CellStyles);

        //Save style for finish
        NewWorkbookStylesPartCreated.Stylesheet = WorkbookStyleSheet;
        NewWorkbookStylesPartCreated.Stylesheet.Save();
    }

    /// <summary>
    /// Gets the SharedStringTablePart of the specified workbook part. If it does not exist, creates a new one.
    /// </summary>
    /// <param name="workbookpart">The workbook part.</param>
    /// <returns>The SharedStringTablePart.</returns>
    internal SharedStringTablePart GetSharedStringTablePart(WorkbookPart workbookpart) => workbookpart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() ?? workbookpart.AddNewPart<SharedStringTablePart>();

    /// <summary>
    /// Inserts a new worksheet into the specified workbook part.
    /// </summary>
    /// <param name="workbookpart">The workbook part.</param>
    /// <param name="sheetname">The name of the new sheet.</param>
    /// <param name="sheetid">The ID of the new sheet.</param>
    /// <returns>The newly created WorksheetPart.</returns>
    internal WorksheetPart InsertWorksheet(WorkbookPart workbookpart, string? sheetname, out string sheetid)
    {
        // Add a new worksheet part to the workbook.
        WorksheetPart NewWorksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        NewWorksheetPart.Worksheet = new Worksheet(new SheetData());
        NewWorksheetPart.Worksheet.Save();

        Sheets sheets = workbookpart.Workbook.GetFirstChild<Sheets>() ?? workbookpart.Workbook.AppendChild(new Sheets());
        string RelationshipId = workbookpart.GetIdOfPart(NewWorksheetPart);
        sheetid = RelationshipId;

        // Get a unique ID for the new sheet.
        uint SheetId = sheets.Elements<Sheet>().Any() ? sheets.Elements<Sheet>().Max(s => s.SheetId!.Value) + 1 : 1;

        string SheetName = sheetname is not null && sheets.Elements<Sheet>().Any(x => x.Name == sheetname) ? sheetname + SheetId : sheetname ?? "Sheet" + SheetId;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = RelationshipId, SheetId = SheetId, Name = SheetName };
        sheets.Append(sheet);
        workbookpart.Workbook.Save();

        return NewWorksheetPart;
    }

    /// <summary>
    /// Inserts a shared string item into the SharedStringTablePart.
    /// </summary>
    /// <param name="text">The text to insert.</param>
    /// <param name="sharestringpart">The SharedStringTablePart.</param>
    /// <param name="sharedstringdictionary">A dictionary to improve lookup performance.</param>
    /// <returns>The index of the inserted shared string item.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal int InsertSharedStringItem(string text, SharedStringTablePart sharestringpart, Dictionary<string, int> sharedstringdictionary)
    {
        //If the part does not contain a SharedStringTable, create one.
        sharestringpart.SharedStringTable ??= new SharedStringTable();

        //Use a dictionary to improve lookup performance.
        if (sharedstringdictionary.TryGetValue(text, out int index))
        {
            return index;
        }

        //The text does not exist in the part. Create the SharedStringItem and return its index.
        index = sharedstringdictionary.Count;

        sharedstringdictionary[text] = index;

        sharestringpart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

        return index;
    }

    /// <summary>
    /// Converts a column index to its corresponding Excel column letters.
    /// </summary>
    /// <param name="columnindex">The column index (0-based).</param>
    /// <returns>The column letters.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the column index is less than zero.</exception>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal string GetColumnLetters(int columnindex)
    {
        if (columnindex < 0)
            throw new ArgumentOutOfRangeException(nameof(columnindex), "Index must be a positive number.");

        const int Base = 26;
        const int ASCIIOffset = 64; // 'A' is 65 in ASCII

        Span<char> ColumnName = stackalloc char[7]; // Max length for Excel columns is "XFD" (3 characters), 7 is more than enough
        int CurrentIndex = ColumnName.Length;

        do
        {
            int Remainder = columnindex % Base;
            ColumnName[--CurrentIndex] = (char)(Remainder + ASCIIOffset + 1);
            //columnindex = (columnindex - Remainder - 1) / Base;
            columnindex = (columnindex / Base) - 1;
        }
        while (columnindex >= 0);

        return new string(ColumnName.Slice(CurrentIndex));
    }

    /// <summary>
    /// Gets the column index from the cell reference.
    /// </summary>
    /// <param name="cellreference">The cell reference (e.g., "A1").</param>
    /// <returns>The column index (1-based).</returns>
    /// <exception cref="ArgumentException">Thrown when the cell reference format is invalid.</exception>
    internal int GetColumnIndex(string cellreference)
    {
        // Use a more efficient method to parse the cell reference without regex.
        int ColumnIndex = 0;
        int i = 0;

        // Process the column letters part.
        do
        {
            ColumnIndex = (ColumnIndex * 26) + (cellreference[i] - 'A' + 1);
            i++;
        }
        while (i < cellreference.Length && char.IsLetter(cellreference[i]));

        if (ColumnIndex <= 0)
        {
            throw new ArgumentException("Invalid cell reference format.");
        }

        return ColumnIndex;
    }
}