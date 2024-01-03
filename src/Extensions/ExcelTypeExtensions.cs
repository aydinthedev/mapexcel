using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class ExcelTypeExtensions
{
    /// <summary>
    ///     Returns true if the ExcelType has worksheet captions.
    /// </summary>
    internal static bool HasWorksheetCaptions(this ExcelType excelType) =>
        excelType.WorksheetCaptions.Count > 0;

    /// <summary>
    ///     Returns the starting index of the first expected worksheet caption row.
    /// </summary>
    internal static int? FirstWorksheetCaptionRow(this ExcelType excelType) =>
        excelType.WorksheetCaptions.Any()
            ? excelType.WorksheetCaptions.Select(h => h.AddressRange).Min(r => r.FirstCellAddress.RowNumber)
            : null;

    /// <summary>
    ///     Returns the starting index of the last expected worksheet caption row.
    /// </summary>
    internal static int? LastWorksheetCaptionRow(this ExcelType excelType) =>
        excelType.WorksheetCaptions.Any()
            ? excelType.WorksheetCaptions.Select(h => h.AddressRange).Max(r => r.LastCellAddress.RowNumber)
            : null;

    /// <summary>
    ///     Returns range between first and last expected worksheet caption rows.
    /// </summary>
    internal static int? WorksheetCaptionRange(this ExcelType excelType) =>
        excelType.LastWorksheetCaptionRow() - excelType.FirstWorksheetCaptionRow() + 1;

    /// <summary>
    ///     Returns true if the ExcelType has column headers.
    /// </summary>
    internal static bool HasColumnHeaders(this ExcelType excelType) =>
        excelType.ColumnHeaderCount > 0;

    /// <summary>
    ///     Returns the starting index of the first expected column header row.
    /// </summary>
    internal static int? FirstColumnHeaderRow(this ExcelType excelType) =>
        excelType.HasColumnHeaders()
            ? (excelType.LastWorksheetCaptionRow() ?? 0) + 1
            : null;

    /// <summary>
    ///     Returns the starting index of the last expected column header row.
    /// </summary>
    internal static int? LastColumnHeaderRow(this ExcelType excelType)
    {
        var columnHeaderStart = excelType.FirstColumnHeaderRow();
        return columnHeaderStart.HasValue
            ? columnHeaderStart + excelType.ColumnHeaderRange() - 1
            : null;
    }

    /// <summary>
    ///     Returns the starting index of the first expected column header column.
    /// </summary>
    internal static int? FirstColumnHeaderColumn(this ExcelType excelType) =>
        excelType.HasColumnHeaders()
            ? excelType.Properties.Min(p => p.ColumnNumber)
            : null;

    /// <summary>
    ///     Returns the starting index of the last expected column header column.
    /// </summary>
    internal static int? LastColumnHeaderColumn(this ExcelType excelType) =>
        excelType.HasColumnHeaders()
            ? excelType.Properties.Max(p => p.ColumnNumber)
            : null;

    /// <summary>
    ///     Returns range between first and last expected column header rows.
    /// </summary>
    internal static int? ColumnHeaderRange(this ExcelType excelType) =>
        excelType.ColumnHeaderCount;

    /// <summary>
    ///     Returns the starting index of the first expected data row.
    /// </summary>
    internal static int FirstDataRow(this ExcelType excelType) =>
        (excelType.LastColumnHeaderRow() ?? 0) + 1;
}