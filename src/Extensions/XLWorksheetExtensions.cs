using ClosedXML.Excel;
using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class XLWorksheetExtensions
{
    /// <summary>
    ///     Returns IXLRange from given CellAddressRange
    /// </summary>
    internal static IXLRange Range(this IXLWorksheet worksheet, CellAddressRange range) =>
        worksheet.Range(range.FirstCellAddress.RowNumber,
            range.FirstCellAddress.ColumnNumber,
            range.LastCellAddress.RowNumber,
            range.LastCellAddress.ColumnNumber);

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
    ///     Returns the first row that contains data:
    ///     - If worksheet is empty, returns null
    ///     - If the worksheet has headers, returns the row below the found header row.
    ///     - If the worksheet has no headers but has captions, returns the row below the found caption row.
    /// </summary>
    internal static IXLRow? FirstDataRowFound(this IXLWorksheet worksheet, WorksheetMetadata metadata)
    {
        if (worksheet.IsEmpty()) return null;

        // Headers always below captions so if headers are expected, we can skip captions
        if (metadata.ExcelType.HasHeaders())
        {
            var foundExpectedHeader = metadata.ExpectedHeaders.FirstOrDefault(x => x.FoundAt != null);
            return foundExpectedHeader != null
                ? worksheet.Row(foundExpectedHeader.FoundAt!.Value.LastCellAddress.RowNumber).RowBelow()
                : null;
        }

        // No headers, no captions so we should use first row
        if (!metadata.ExcelType.HasCaptions()) return worksheet.FirstRowUsed();

        var foundExpectedCaption = metadata.ExpectedCaptions.FirstOrDefault(x => x.FoundAt != null);
        return foundExpectedCaption != null
            ? worksheet.Row(foundExpectedCaption.FoundAt!.Value.LastCellAddress.RowNumber).RowBelow()
            : null;
    }
}