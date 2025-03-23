namespace CsvHelper.OpenXml.Excel.Tests.ConsoleApp;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public enum VerificationEnumeration : int
{
    All = 0,
    DomImportExcelFile = 1,
    DomImportExcelFileAsyncEnumerable = 2,
    DomImportExcelFileDynamic = 3,
    DomImportExcelFileDynamicAsyncEnumerable = 4,
    SaxImportExcelFile = 5,
    SaxImportExcelFileAsyncEnumerable = 6,
    SaxImportExcelFileDynamic = 7,
    SaxImportExcelFileDynamicAsyncEnumerable = 8,
    DomExportExcelFile = 9,
    DomExportExcelFileTwoCollectionSameTypeOnSameSheet = 10,
    DomExportExcelFileTwoCollectionSameTypeOnDifferentSheet = 11,
    DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheet = 12,
    DomExportExcelFileAsync = 13,
    DomExportExcelFileTwoCollectionSameTypeOnSameSheetAsync = 14,
    DomExportExcelFileTwoCollectionSameTypeOnDifferentSheetAsync = 15,
    DomExportExcelFileTwoCollectionDifferentTypeOnDifferentSheetAsync = 16,

}