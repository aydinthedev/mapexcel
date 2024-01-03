using ClosedXML.Excel;
using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class XLRowExtensions
{
    /// <summary>
    ///     Returns the cell in the given row that matches the given excelProperty
    ///     and the excelProperty match strategy.
    /// </summary>
    internal static IXLCell? Cell(this IXLRow row, WorksheetMetadata metadata, ExcelProperty excelProperty) =>
        excelProperty.ColumnMatch switch
        {
            ColumnMatch.ByColumn => row.Cell(excelProperty.ColumnNumber),
            ColumnMatch.ByHeader =>
                metadata.FoundHeaders.TryGetValue(excelProperty.HeaderName, out var headerMetadata)
                    ? row.Cell(headerMetadata.FoundAt!.Value.ColumnNumber)
                    : null,
            ColumnMatch.ByColumnAndHeader =>
                metadata.FoundHeaders.TryGetValue(excelProperty.HeaderName, out var headerMetadata)
                    ? headerMetadata.FoundAt!.Value.ColumnNumber == excelProperty.ColumnNumber
                        ? row.Cell(excelProperty.ColumnNumber)
                        : null
                    : null,
            _ => throw new ArgumentOutOfRangeException(nameof(excelProperty))
        };
}