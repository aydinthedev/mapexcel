using System.Reflection;
using ClosedXML.Excel;
using MapExcel.Exceptions;
using MapExcel.Extensions;

namespace MapExcel.Metadata;

public class WorksheetMetadata
{
    private readonly ExcelType _excelType;
    private readonly IXLWorksheet _worksheet;

    internal WorksheetMetadata(IXLWorksheet worksheet, ExcelType excelType, bool isNewWorksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(excelType);

        _worksheet = worksheet;
        _excelType = excelType;

        Name = worksheet.Name;
        IsNew = isNewWorksheet;
        ExpectedFirstDataRowNumber = excelType.FirstDataRow();

        // Extract expected captions from ExcelType
        Captions = _excelType.WorksheetCaptions
            .Select(caption => new WorksheetCaptionMetadata
            {
                Name = caption.Name,
                ExpectedAtRange = caption.AddressRange
            }).ToList();

        // Extract expected headers from ExcelType
        var firstColumnHeaderRow = excelType.FirstColumnHeaderRow();
        if (firstColumnHeaderRow != null)
        {
            MissingHeaders = ExpectedHeaders = excelType.Properties
                .ToDictionary(
                    property => property.HeaderName,
                    property => new ColumnHeaderMetadata
                    {
                        PropertyInfo = property.PropertyInfo,
                        ExpectedAt = new CellAddress(firstColumnHeaderRow.Value, property.ColumnNumber)
                    });

            FoundHeaders = new Dictionary<string, ColumnHeaderMetadata>();
        }
        else
        {
            MissingHeaders = ExpectedHeaders = FoundHeaders = new Dictionary<string, ColumnHeaderMetadata>();
        }

        UpdateMetadata(worksheet, excelType);
    }

    public string Name { get; }

    public bool IsNew { get; }

    public bool IsEmpty { get; private set; }

    public int ExpectedFirstDataRowNumber { get; }

    public int? FoundFirstDataRowNumber { get; private set; }

    public IReadOnlyList<WorksheetCaptionMetadata> Captions { get; }

    public IReadOnlyDictionary<string, ColumnHeaderMetadata> ExpectedHeaders { get; }

    public IReadOnlyDictionary<string, ColumnHeaderMetadata> FoundHeaders { get; private set; }

    public IReadOnlyDictionary<string, ColumnHeaderMetadata> MissingHeaders { get; private set; }

    internal void UpdateWorksheetCaptionMetadata()
    {
        foreach (var caption in Captions)
        {
            var rowNumber = caption.ExpectedAtRange.FirstCellAddress.RowNumber;
            var columnNumber = caption.ExpectedAtRange.FirstCellAddress.ColumnNumber;

            var cell = _worksheet.Cell(rowNumber, columnNumber);

            caption.IsFound = caption.Name == cell.GetString().Trim();
        }
    }

    internal void UpdateColumnHeaderMetadata()
    {
        var foundHeaderStartRowNumber = _worksheet.FirstColumnHeaderRow(_excelType);
        if (foundHeaderStartRowNumber == null)
            return;

        var foundHeaders = new Dictionary<string, ColumnHeaderMetadata>();

        var headerCells = _worksheet
            .Row(foundHeaderStartRowNumber.Value)
            .CellsUsed(cell => cell.DataType != XLDataType.Blank);

        foreach (var cell in headerCells)
        {
            var header = cell.GetString().Trim();

            var cellAddress = new CellAddress(cell.Address.RowNumber, cell.Address.ColumnNumber);

            if (foundHeaders.TryGetValue(header, out var duplicateHeader))
                throw new DuplicateHeaderException(
                    _worksheet,
                    duplicateHeader.FoundAt!.Value,
                    cellAddress,
                    header);

            var headerMetadata = ExpectedHeaders.TryGetValue(header, out var value)
                ? value
                : new ColumnHeaderMetadata();

            headerMetadata.FoundAt = cellAddress;

            foundHeaders.Add(header, headerMetadata);
        }

        FoundHeaders = foundHeaders;

        MissingHeaders = ExpectedHeaders.Except(FoundHeaders).ToDictionary(pair => pair.Key, pair => pair.Value);
    }

    internal void UpdateMetadata(IXLWorksheet worksheet, ExcelType excelType)
    {
        IsEmpty = worksheet.IsEmpty();

        UpdateWorksheetCaptionMetadata();

        UpdateColumnHeaderMetadata();

        FoundFirstDataRowNumber = worksheet.FirstDataRow(excelType);
    }
}

public class WorksheetCaptionMetadata
{
    // Prevent outside initialization
    internal WorksheetCaptionMetadata()
    {
    }

    public string? Name { get; internal set; }

    public CellAddressRange ExpectedAtRange { get; internal set; }

    public bool IsFound { get; internal set; }
}

public class ColumnHeaderMetadata
{
    // Prevent outside initialization
    internal ColumnHeaderMetadata()
    {
    }

    public PropertyInfo? PropertyInfo { get; internal set; }

    public CellAddress? ExpectedAt { get; internal set; }

    public CellAddress? FoundAt { get; internal set; }
}