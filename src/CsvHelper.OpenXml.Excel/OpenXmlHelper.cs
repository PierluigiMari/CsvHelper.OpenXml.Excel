using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("CsvHelper.OpenXml.Excel.Tests")]

namespace CsvHelper.OpenXml.Excel;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

internal class OpenXmlHelper
{
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
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 170, FormatCode = "#,##0.00\\ \"€\"" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 171, FormatCode = "#,##0.0000\\ \"€\"" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 172, FormatCode = "_-* #,##0.0000 \"€\"_-;-* #,##0.0000 \"€\"_-;_-* \"-\"???? \"€\"_-;_-@_-" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 173, FormatCode = "[$-F400]h:mm:ss\\ AM/PM" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 174, FormatCode = "h:mm;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 175, FormatCode = "[$-409]h:mm:ss\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 176, FormatCode = "[$-409]h:mm\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 177, FormatCode = "dd/mm/yyyy\\ hh:mm:ss" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 178, FormatCode = "dd/mm/yyyy\\ h:mm:ss\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 179, FormatCode = "dd/mm/yyyy\\ h:mm\\ AM/PM;@" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 180, FormatCode = "0.0000E+00" });
        StylesheetNumberingFormats.Append(new NumberingFormat { NumberFormatId = 181, FormatCode = "00000" });

        Font FontDefault = new Font(new FontName { Val = "Calibri" }, new FontSize { Val = 11 }); // Default font
        Font FontBold = new Font(new Bold()); // Bold font

        Fonts Fonts = new Fonts(FontDefault, FontBold);

        Fill FillDefault = new Fill(new PatternFill { PatternType = PatternValues.None }); // Default fill
        Fills Fills = new Fills(FillDefault);

        Border BorderDefault = new Border(); // Default border
        Borders Borders = new Borders(BorderDefault);

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
        CellFormat CurrencyEuroFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 170, FormatId = 0, ApplyNumberFormat = true }; // format like "#,##0.00 €"
        CellFormat CurrencyEuroFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 171, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 44, FormatId = 0, ApplyNumberFormat = true };
        CellFormat AccountingFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 172, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateFormatDefault = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 14, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateFormatExtended = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 165, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateFormatWithDash = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 166, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTimeFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 177, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTimeFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 22, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTime12HourFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 178, FormatId = 0, ApplyNumberFormat = true };
        CellFormat DateTime12HourFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 179, FormatId = 0, ApplyNumberFormat = true };
        CellFormat TimeFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 173, FormatId = 0, ApplyNumberFormat = true };
        CellFormat TimeFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 174, FormatId = 0, ApplyNumberFormat = true };
        CellFormat Time12HourFormatWithHoursMinutesSeconds = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 175, FormatId = 0, ApplyNumberFormat = true };
        CellFormat Time12HourFormatWithHoursMinutes = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 176, FormatId = 0, ApplyNumberFormat = true };
        CellFormat PercentageFormatWithoutDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 9, FormatId = 0, ApplyNumberFormat = true };
        CellFormat PercentageFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 10, FormatId = 0, ApplyNumberFormat = true };
        CellFormat ScientificFormatWithTwoDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 11, FormatId = 0, ApplyNumberFormat = true };
        CellFormat ScientificFormatWithFourDecimals = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 180, FormatId = 0, ApplyNumberFormat = true };
        CellFormat SpecialZipCodeFormat = new CellFormat { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 181, FormatId = 0, ApplyNumberFormat = true };

        CellFormats.Append(CellFormatDefault);
        CellFormats.Append(CellFormatDefaultBold);
        CellFormats.Append(CellFormatDefaultBoldCentered);
        CellFormats.Append(NumberIntegerFormat);
        CellFormats.Append(NumberDecimalFormatWithTwoDecimals);
        CellFormats.Append(NumberDecimalFormatWithFourDecimals);
        CellFormats.Append(CurrencyGenericFormatWithoutDecimals);
        CellFormats.Append(CurrencyGenericFormatWithTwoDecimals);
        CellFormats.Append(CurrencyGenericFormatWithFourDecimals);
        CellFormats.Append(CurrencyEuroFormatWithTwoDecimals);
        CellFormats.Append(CurrencyEuroFormatWithFourDecimals);
        CellFormats.Append(AccountingFormatWithTwoDecimals);
        CellFormats.Append(AccountingFormatWithFourDecimals);
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
        CellFormats.Append(SpecialZipCodeFormat);

        // Append everything to stylesheet  - Preserve the ORDER!
        WorkbookStyleSheet.Append(StylesheetNumberingFormats);
        WorkbookStyleSheet.Append(Fonts);
        WorkbookStyleSheet.Append(Fills);
        WorkbookStyleSheet.Append(Borders);
        WorkbookStyleSheet.Append(CellFormats);

        //Save style for finish
        NewWorkbookStylesPartCreated.Stylesheet = WorkbookStyleSheet;
        NewWorkbookStylesPartCreated.Stylesheet.Save();
    }

    internal SharedStringTablePart GetSharedStringTablePart(WorkbookPart workbookpart) => workbookpart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() ?? workbookpart.AddNewPart<SharedStringTablePart>();

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

    internal int InsertSharedStringItem(string text, SharedStringTablePart sharestringpart)
    {
        // If the part does not contain a SharedStringTable, create one.
        sharestringpart.SharedStringTable ??= new SharedStringTable();

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in sharestringpart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
                return i;

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        sharestringpart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
        sharestringpart.SharedStringTable.Save();

        return i;
    }

    internal string GetColumnLetters(int columnindex)
    {
        columnindex++;
        const int Base = 26;
        const int ASCIIOffset = 64; // 'A' is 65 in ASCII

        if (columnindex <= 0)
            throw new ArgumentOutOfRangeException(nameof(columnindex), "Index must be a positive number.");

        StringBuilder ColumnNameBuilder = new StringBuilder();

        while (columnindex > 0)
        {
            int Remainder = (columnindex - 1) % Base;
            ColumnNameBuilder.Insert(0, (char)(Remainder + ASCIIOffset + 1));
            columnindex = (columnindex - Remainder - 1) / Base;
        }

        return ColumnNameBuilder.ToString();
    }

    internal int GetColumnIndex(string cellreference)
    {
        Match RegexMatch = Regex.Match(cellreference, @"([A-Za-z]+)(\d+)");

        if (!RegexMatch.Success)
        {
            throw new ArgumentException("Invalid cell reference format.");
        }

        string ColumnLettersPart = RegexMatch.Groups[1].Value;

        return ColumnLettersPart.Aggregate(0, (index, letter) => (index * 26) + (letter - 'A' + 1));
    }
}