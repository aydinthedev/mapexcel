using ClosedXML.Excel;
using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class XLWorksheetExtensions
{
    /// <summary>
    ///     Returns the row below the last non-null row
    /// </summary>
    internal static IXLRow NextEmptyRow(this IXLWorksheet worksheet)
    {
        // When a row is merged, LastRowUsed() returns first row of the merged range
        // So we need to find the last row in the merged range to get the next empty row
        var lastUsedRow = worksheet.LastRowUsed();
        return lastUsedRow == null
            ? worksheet.FirstRow()
            : lastUsedRow.IsMerged()
                ? worksheet.MergedRanges.Last().LastRow().RowBelow().WorksheetRow()
                : lastUsedRow.RowBelow();
    }

    /// <summary>
    ///     Returns the starting index of the first found worksheet caption row.
    /// </summary>
    internal static int? FirstWorksheetCaptionRow(this IXLWorksheet worksheet, ExcelType excelType) =>
        worksheet.IsEmpty() || !excelType.HasWorksheetCaptions()
            ? null
            : worksheet.FirstRowUsed().RowNumber();

    /// <summary>
    ///     Returns the starting index of the last found worksheet caption row.
    /// </summary>
    internal static int? LastWorksheetCaptionRow(this IXLWorksheet worksheet, ExcelType excelType)
    {
        var firstWorksheetCaptionRow = worksheet.FirstWorksheetCaptionRow(excelType);
        if (firstWorksheetCaptionRow == null)
            return null;

        return firstWorksheetCaptionRow + excelType.WorksheetCaptionRange() - 1;
    }

    /// <summary>
    ///     Returns the starting index of the first found column header row.
    /// </summary>
    internal static int? FirstColumnHeaderRow(this IXLWorksheet worksheet, ExcelType excelType)
    {
        if (worksheet.IsEmpty() || !excelType.HasColumnHeaders())
            return null;

        var lastWorksheetCaptionRow = worksheet.LastWorksheetCaptionRow(excelType);
        return lastWorksheetCaptionRow != null
            ? lastWorksheetCaptionRow + 1
            : worksheet.FirstRowUsed().RowNumber();
    }

    /// <summary>
    ///     Returns the starting index of the last found column header row.
    /// </summary>
    internal static int? LastColumnHeaderRow(this IXLWorksheet worksheet, ExcelType excelType)
    {
        var firstColumnHeaderRow = worksheet.FirstColumnHeaderRow(excelType);
        if (firstColumnHeaderRow == null)
            return null;

        return firstColumnHeaderRow + excelType.ColumnHeaderRange() - 1;
    }

    /// <summary>
    ///     Returns the starting index of the first found data row.
    /// </summary>
    internal static int? FirstDataRow(this IXLWorksheet worksheet, ExcelType excelType)
    {
        if (worksheet.IsEmpty())
            return null;

        var lastColumnHeaderRow = worksheet.LastColumnHeaderRow(excelType);
        return lastColumnHeaderRow != null
            ? lastColumnHeaderRow + 1
            : worksheet.FirstRowUsed().RowNumber();
    }
}