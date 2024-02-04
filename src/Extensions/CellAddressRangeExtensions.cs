namespace MapExcel.Extensions;

internal static class CellAddressRangeExtensions
{
    internal static int RowSpan(this CellAddressRange range) =>
        range.LastCellAddress.RowNumber - range.FirstCellAddress.RowNumber;

    internal static int ColumnSpan(this CellAddressRange range) =>
        range.LastCellAddress.ColumnNumber - range.FirstCellAddress.ColumnNumber;
}