using ClosedXML.Excel;
using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class XLWorkbookExtensions
{
    /// <summary>
    ///     Returns the worksheet that matches with the worksheet strategy
    /// </summary>
    internal static IXLWorksheet? Worksheet(this IXLWorkbook workbook, ExcelType excelType) =>
        excelType.WorksheetMatch switch
        {
            WorksheetMatch.ByNumber =>
                workbook.Worksheets.Count >= excelType.WorksheetNumber
                    ? workbook.Worksheet(excelType.WorksheetNumber)
                    : null,
            WorksheetMatch.ByName =>
                workbook.TryGetWorksheet(excelType.WorksheetName, out var byName)
                    ? byName
                    : null,
            WorksheetMatch.ByNumberAndName =>
                workbook.TryGetWorksheet(excelType.WorksheetName, out var byIndexAndName)
                    ? byIndexAndName.Position == excelType.WorksheetNumber
                        ? byIndexAndName
                        : null
                    : null,
            WorksheetMatch.PreferNumber =>
                workbook.Worksheets.Count >= excelType.WorksheetNumber
                    ? workbook.Worksheet(excelType.WorksheetNumber)
                    : workbook.TryGetWorksheet(excelType.WorksheetName, out var preferIndex)
                        ? preferIndex
                        : null,
            WorksheetMatch.PreferName =>
                workbook.TryGetWorksheet(excelType.WorksheetName, out var preferName)
                    ? preferName
                    : workbook.Worksheets.Count >= excelType.WorksheetNumber
                        ? workbook.Worksheet(excelType.WorksheetNumber)
                        : null,
            _ => throw new ArgumentOutOfRangeException(nameof(excelType))
        };

    /// <summary>
    ///     Returns the worksheet that matches with the worksheet strategy to update
    ///     or if not found creates a new one.
    /// </summary>
    internal static (bool IsNewWorksheet, IXLWorksheet Worksheet)? GetOrAddWorksheet(
        this IXLWorkbook workbook, ExcelType excelType)
    {
        // When there is a sheet that matches the worksheet strategy, return it
        // If we can read we should be able to write
        var readableSheet = workbook.Worksheet(excelType);
        if (readableSheet != null)
            return (false, readableSheet);

        // We fall here if there is no exact number or name match
        // If we have missing sheets then those are the reason we can't find a match
        // Add the missing sheets and return the last one
        var missingSheets = excelType.WorksheetNumber - workbook.Worksheets.Count;
        if (missingSheets > 0)
        {
            // Skip the last sheet; it will be added later
            for (var i = 0; i < missingSheets - 1; i++)
                workbook.AddWorksheet();

            var lastSheet = workbook.AddWorksheet();
            return (true, lastSheet);
        }

        // Here we know we have enough sheets, but we couldn't find a match
        // So there is no sheet with the specified name
        // or there is one but it has a different number
        // If the sheet at the specified number is empty, we can use it
        var sameIndexSheet = workbook.Worksheet(excelType.WorksheetNumber);
        if (sameIndexSheet.IsEmpty())
            return (false, sameIndexSheet);

        // Since the sheet at the specified number is not empty and does not match the name
        // we can't be sure it belongs to this type, so we can't use it
        // In this case last thing we can do is to create a new sheet with the specified name
        // if the worksheet strategy is ByName
        // Because of we match by name, we actually don't care about the number
        if (excelType.WorksheetMatch != WorksheetMatch.ByName)
            return null;

        var createdSheet = workbook.AddWorksheet();

        return (true, createdSheet);
    }
}