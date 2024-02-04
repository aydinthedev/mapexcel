using ClosedXML.Excel;
using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class XLRowExtensions
{
    /// <summary>
    ///     Returns the cell at the given row that matches the given ExcelProperty.
    /// </summary>
    internal static IXLCell Cell(this IXLRow row, WorksheetMetadata metadata, ExcelProperty excelProperty) =>
        // If headers not enabled use the column number directly
        metadata.PropertyHeaderMap.TryGetValue(excelProperty, out var headerMetadata)
            ? row.Cell(headerMetadata!.FoundAt!.Value.FirstCellAddress.ColumnNumber)
            : row.Cell(excelProperty.ColumnNumber);
}