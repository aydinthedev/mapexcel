using ClosedXML.Excel;
using MapExcel.Exceptions;
using MapExcel.Extensions;

namespace MapExcel.Metadata;

public class WorksheetMetadata
{
    private readonly List<CaptionMetadata> _expectedCaptions = new();
    private readonly List<HeaderMetadata> _expectedHeaders = new();
    private readonly List<CaptionMetadata> _foundCaptions = new();
    private readonly List<HeaderMetadata> _foundHeaders = new();
    private readonly List<CaptionMetadata> _missingCaptions = new();
    private readonly List<HeaderMetadata> _missingHeaders = new();
    private readonly Dictionary<ExcelProperty, HeaderMetadata?> _propertyHeaderMap = new();

    internal WorksheetMetadata(IXLWorksheet worksheet, ExcelType excelType, bool isNewWorksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(excelType);

        Worksheet = worksheet;
        ExcelType = excelType;
        IsNew = isNewWorksheet;

        Initialize();
        Update();
    }

    internal IXLWorksheet Worksheet { get; }

    internal ExcelType ExcelType { get; }

    public bool IsNew { get; }

    public bool IsEmpty => Worksheet.IsEmpty();

    public string Name => Worksheet.Name;

    public IReadOnlyList<CaptionMetadata> ExpectedCaptions => _expectedCaptions;

    public IReadOnlyList<CaptionMetadata> FoundCaptions => _foundCaptions;

    public IReadOnlyList<CaptionMetadata> MissingCaptions => _missingCaptions;

    public IReadOnlyList<HeaderMetadata> ExpectedHeaders => _expectedHeaders;

    public IReadOnlyList<HeaderMetadata> FoundHeaders => _foundHeaders;

    public IReadOnlyList<HeaderMetadata> MissingHeaders => _missingHeaders;

    internal IReadOnlyDictionary<ExcelProperty, HeaderMetadata?> PropertyHeaderMap => _propertyHeaderMap;

    public bool Validate(bool checkCaptionPosition = false, bool ignoreExtraHeaders = true) =>
        ValidateCaptions(checkCaptionPosition) && ValidateHeaders(ignoreExtraHeaders);

    public bool ValidateCaptions(bool checkPosition = false) =>
        IsEmpty || (MissingCaptions.Count == 0 && ExpectedCaptions.All(x => x.Validate(checkPosition)));

    public bool ValidateHeaders(bool ignoreExtraHeaders = true)
    {
        if (IsEmpty) return true;
        var hasNoInvalidHeader = MissingHeaders.Count == 0 && ExpectedHeaders.All(x => x.Validate());
        return hasNoInvalidHeader && (ignoreExtraHeaders || FoundHeaders.Count == ExpectedHeaders.Count);
    }

    internal void Update()
    {
        UpdateCaptions();
        UpdateHeaders();
    }

    private void Initialize()
    {
        if (ExcelType.HasCaptions())
        {
            var expectedCaptions = ExcelType.Captions
                .Select(x => new CaptionMetadata
                {
                    Name = x.Name,
                    ExpectedAt = x.AddressRange
                });

            _expectedCaptions.AddRange(expectedCaptions);
        }

        if (!ExcelType.HasHeaders())
            return;

        // Find the row below last expected caption
        // If no captions are expected, then the header row is the first row
        var firstHeaderRowNumberExpected = ExcelType.HasCaptions()
            ? ExcelType.Captions.Max(x => x.AddressRange.LastCellAddress).RowNumber + 1
            : 1;

        var headerRowSpan = ExcelType.HeaderRows - 1;

        // The header row can span multiple rows, so we need to calculate the last header row number
        var lastHeaderRowNumberExpected = firstHeaderRowNumberExpected + headerRowSpan;

        // Extract expected headers from ExcelType
        var expectedHeaders = ExcelType.Properties
            .Select(x => new HeaderMetadata(x.HeaderName)
            {
                ExcelProperty = x,
                ExpectedAt = new CellAddressRange(
                    firstHeaderRowNumberExpected,
                    x.ColumnNumber,
                    lastHeaderRowNumberExpected,
                    x.ColumnNumber)
            });

        foreach (var expectedHeader in expectedHeaders)
        {
            _expectedHeaders.Add(expectedHeader);
            _propertyHeaderMap.Add(expectedHeader.ExcelProperty!, expectedHeader);
        }
    }

    private void UpdateCaptions()
    {
        if (!ExcelType.HasCaptions())
            return;

        _foundCaptions.Clear();
        _missingCaptions.Clear();

        // If the worksheet is empty, then there is no need to search for captions
        var firstCellUsed = Worksheet.FirstRowUsed()?.FirstCellUsed();
        if (firstCellUsed == null)
        {
            _missingCaptions.AddRange(_expectedCaptions);
            return;
        }

        // Find the offset between the first cell used and the first caption start position
        var firstCaptionAddress = ExcelType.Captions.Min(x => x.AddressRange.FirstCellAddress);
        var rowOffset = firstCellUsed.Address.RowNumber - firstCaptionAddress.RowNumber;
        var columnOffset = firstCellUsed.Address.ColumnNumber - firstCaptionAddress.ColumnNumber;

        foreach (var expectedCaption in _expectedCaptions)
        {
            var firstRowNumber = expectedCaption.ExpectedAt.FirstCellAddress.RowNumber + rowOffset;
            var firstColumnNumber = expectedCaption.ExpectedAt.FirstCellAddress.ColumnNumber + columnOffset;

            var foundRange = new CellAddressRange(
                firstRowNumber,
                firstColumnNumber,
                firstRowNumber + expectedCaption.ExpectedAt.RowSpan(),
                firstColumnNumber + expectedCaption.ExpectedAt.ColumnSpan());

            var worksheetRange = Worksheet.Range(foundRange);

            if (worksheetRange.FirstCellUsed()?.GetString() != expectedCaption.Name) continue;

            expectedCaption.FoundAt = foundRange;
            _foundCaptions.Add(expectedCaption);
        }

        _missingCaptions.AddRange(_expectedCaptions.Except(_foundCaptions));
    }

    private void UpdateHeaders()
    {
        if (!ExcelType.HasHeaders())
            return;

        _foundHeaders.Clear();
        _missingHeaders.Clear();

        // If the worksheet is empty, then there is no need to search for headers
        var firstRowUsed = Worksheet.FirstRowUsed();
        if (firstRowUsed == null)
        {
            _missingHeaders.AddRange(_expectedHeaders);
            return;
        }

        // Find the actual header row number in the worksheet
        int headerRowNumberFound;
        if (ExcelType.HasCaptions())
        {
            var lastCaption = _expectedCaptions.MaxBy(x => x.ExpectedAt.LastCellAddress);
            if (lastCaption?.FoundAt == null)
            {
                _missingHeaders.AddRange(_expectedHeaders);
                return;
            }

            var lastCaptionRowNumber = lastCaption.FoundAt.Value.LastCellAddress.RowNumber;

            // The header row is the row below the last caption
            headerRowNumberFound = lastCaptionRowNumber + 1;
        }
        else
        {
            // When there is no caption, the first row used should be the header row
            headerRowNumberFound = firstRowUsed.RowNumber();
        }

        var cells = Worksheet
            .Row(headerRowNumberFound)
            .CellsUsed(cell => cell.DataType != XLDataType.Blank);

        foreach (var cell in cells)
        {
            var header = cell.GetString().Trim();
            var address = new CellAddress(cell.Address.RowNumber, cell.Address.ColumnNumber);

            var duplicateHeader = _foundHeaders.FirstOrDefault(x => x.Name == header);
            if (duplicateHeader != null)
            {
                var duplicateAddress = duplicateHeader.FoundAt!.Value.FirstCellAddress;
                throw new DuplicateHeaderException(Worksheet, address, duplicateAddress, header);
            }

            // If we don't expect this header, then create a new HeaderMetadata
            var headerMetadata = _expectedHeaders.FirstOrDefault(x => x.Name == header) ?? new HeaderMetadata(header);

            var headerRowSpan = ExcelType.HeaderRows - 1;

            // Headers can span multiple rows, requiring us to identify the last row within the merged range
            var foundRange = new CellAddressRange(
                address.RowNumber,
                address.ColumnNumber,
                address.RowNumber + headerRowSpan,
                address.ColumnNumber);

            headerMetadata.FoundAt = foundRange;
            _foundHeaders.Add(headerMetadata);
        }

        _missingHeaders.AddRange(_expectedHeaders.Except(_foundHeaders));
    }
}