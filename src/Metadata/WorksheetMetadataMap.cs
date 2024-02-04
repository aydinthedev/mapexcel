using ClosedXML.Excel;
using MapExcel.Extensions;

namespace MapExcel.Metadata;

internal sealed class WorksheetMetadataMap
{
    private readonly Dictionary<ExcelType, WorksheetMetadata> _maps = new();
    private readonly IXLWorkbook _workbook;

    public WorksheetMetadataMap(IXLWorkbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        _workbook = workbook;
    }

    internal WorksheetMetadata? TryGet(ExcelType excelType)
    {
        if (_maps.TryGetValue(excelType, out var value))
            return value;

        var worksheet = _workbook.Worksheet(excelType);
        if (worksheet == null)
            return null;

        var metadata = new WorksheetMetadata(worksheet, excelType, false);

        _maps[excelType] = metadata;

        return metadata;
    }

    internal WorksheetMetadata GetOrAdd(ExcelType excelType)
    {
        if (_maps.TryGetValue(excelType, out var value))
            return value;

        var (worksheet, isNew) =
            _workbook.GetOrAddWorksheet(excelType)
            ?? throw new InvalidOperationException(
                "There is another worksheet in the same same position.");

        var metadata = new WorksheetMetadata(worksheet, excelType, isNew);

        _maps[excelType] = metadata;

        return metadata;
    }
}