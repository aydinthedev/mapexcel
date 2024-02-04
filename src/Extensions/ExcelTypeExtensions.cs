using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class ExcelTypeExtensions
{
    /// <summary>
    ///     Returns true if the ExcelType has captions.
    /// </summary>
    internal static bool HasCaptions(this ExcelType excelType) =>
        excelType.Captions.Count > 0;

    /// <summary>
    ///     Returns true if the ExcelType has headers.
    /// </summary>
    internal static bool HasHeaders(this ExcelType excelType) =>
        excelType.HeaderRows > 0;
}