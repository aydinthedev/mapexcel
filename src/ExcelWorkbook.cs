using ClosedXML.Excel;
using MapExcel.Exceptions;
using MapExcel.Extensions;
using MapExcel.Metadata;

namespace MapExcel;

public sealed class ExcelWorkbook : IDisposable
{
    private readonly Dictionary<ExcelType, (IXLWorksheet Worksheet, WorksheetMetadata Metadata)> _metadataMap = new();
    private readonly IXLWorkbook _workbook;

    public ExcelWorkbook()
    {
        _workbook = new XLWorkbook();
    }

    public ExcelWorkbook(string filePath)
    {
        _workbook = new XLWorkbook(filePath);
    }

    public ExcelWorkbook(Stream fileStream)
    {
        _workbook = new XLWorkbook(fileStream);
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }

    public WorksheetMetadata? GetMetadata<T>() where T : class =>
        GetMetadataMap(ExcelTypeRegistry.Get(typeof(T)))?.Metadata;

    public IEnumerable<RowData<T>> Read<T>(WorksheetOptions? options = null) where T : class
    {
        var typeMap = ExcelTypeRegistry.Get(typeof(T));

        var (worksheet, metadata) =
            GetMetadataMap(typeMap)
            ?? throw new InvalidOperationException("Associated worksheet is not found");

        if (metadata.IsEmpty)
            yield break;

        var useOptions = options ?? WorksheetOptions.Default;
        foreach (var row in worksheet.RowsUsed(x => x.RowNumber() >= metadata.FoundFirstDataRowNumber))
        {
            var entity = Activator.CreateInstance<T>();
            var rowData = new RowData<T>(entity, row.RowNumber());

            foreach (var property in typeMap.Properties)
            {
                var cell = row.Cell(metadata, property)
                           ?? throw new CellMismatchException(
                               new CellAddress(row.RowNumber(), property.ColumnNumber), property.HeaderName);

                // Use custom deserializer if exists
                object? value;
                if (property.Deserializer != null)
                {
                    var context = new DeserializationContext(useOptions, property.PropertyInfo, cell.Style);
                    value = property.Deserializer.Invoke(cell.Value, context);

                    // User can set error message to context.Error
                    // If its set then conversion is failed so we must add error to row data
                    // and continue to next property
                    if (!context.IsValid)
                    {
                        rowData.InvalidateProperty(property.PropertyInfo);
                        continue;
                    }
                }
                else
                {
                    if (!cell.TryConvertToObject(property, useOptions, out value))
                    {
                        rowData.InvalidateProperty(property.PropertyInfo);
                        continue;
                    }
                }

                // If read value is null, do not set the property.
                // This is to prevent setting default values to properties.
                if (value == null)
                    continue;

                var ignore = property.Ignore?.Invoke(property.PropertyInfo, value, cell.Style, RowEvent.Read);
                if (ignore == true)
                    continue;

                property.PropertyInfo.SetValue(entity, value);
            }

            yield return rowData;
        }
    }

    public void Write<T>(T entity, WorksheetOptions? options = null) where T : class
    {
        ArgumentNullException.ThrowIfNull(entity);

        var excelType = ExcelTypeRegistry.Get<T>();

        var metadataMap = GetOrAddMetadataMap(excelType);

        WriteRow(excelType, metadataMap, entity, options);
    }

    public void Write<T>(IEnumerable<T> entities, WorksheetOptions? options = null) where T : class
    {
        ArgumentNullException.ThrowIfNull(entities);

        var excelType = ExcelTypeRegistry.Get<T>();

        var metadataMap = GetOrAddMetadataMap(excelType);

        foreach (var entity in entities)
            WriteRow(excelType, metadataMap, entity, options);
    }

    public void Save() => _workbook.Save();

    public void SaveAs(string filePath) => _workbook.SaveAs(filePath);

    public void SaveAs(Stream stream) => _workbook.SaveAs(stream);

    private (IXLWorksheet Worksheet, WorksheetMetadata Metadata)? GetMetadataMap(
        ExcelType excelType, bool isNewWorksheet = false)
    {
        if (_metadataMap.TryGetValue(excelType, out var value))
            return value;

        var worksheet = _workbook.Worksheet(excelType);
        if (worksheet == null)
            return null;

        var metadata = new WorksheetMetadata(worksheet, excelType, isNewWorksheet);
        var metadataMap = (worksheet, metadata);
        _metadataMap[excelType] = metadataMap;

        return metadataMap;
    }

    private (IXLWorksheet Worksheet, WorksheetMetadata Metadata) GetOrAddMetadataMap(ExcelType excelType)
    {
        var metadataMap = GetMetadataMap(excelType);

        if (metadataMap != null)
            return metadataMap.Value;

        var newOrExistingWorksheet =
            _workbook.GetOrAddWorksheet(excelType)
            ?? throw new InvalidOperationException("There is another worksheet in the same same position.");

        CreateTemplate(newOrExistingWorksheet.Worksheet, excelType);

        // Create metadata after writing template
        // So metadata can analyze the headers on initialization
        return GetMetadataMap(excelType, newOrExistingWorksheet.IsNewWorksheet)
               ?? throw new InvalidOperationException("Metadata is not found");
    }

    private static void CreateTemplate(IXLWorksheet worksheet, ExcelType excelType)
    {
        // Check if the worksheet is empty. If it is, set the worksheet name and write headers.
        // For new sheets, headers need to be written to ensure accurate data retrieval later,
        // as column positions might be changed by users. For existing sheets, headers should
        // not be overwritten to preserve user-defined column orders.
        if (!worksheet.IsEmpty())
            return;

        worksheet.Name = excelType.WorksheetName;

        CreateHeaders(worksheet, excelType);
    }

    private static void CreateHeaders(IXLWorksheet worksheet, ExcelType excelType)
    {
        foreach (var caption in excelType.WorksheetCaptions)
        {
            var firstCellAddress = caption.AddressRange.FirstCellAddress;
            var lastCellAddress = caption.AddressRange.LastCellAddress;

            var captionRange = worksheet.Range(
                firstCellAddress.RowNumber, firstCellAddress.ColumnNumber,
                lastCellAddress.RowNumber, lastCellAddress.ColumnNumber);

            captionRange.Merge();
            captionRange.SetValue(caption.Name);
            caption.Style?.Invoke(captionRange.Style);
        }

        // If column header row count is 0, skip writing column headers
        if (!excelType.HasColumnHeaders())
            return;

        var columnHeaderRowStart = excelType.FirstColumnHeaderRow()!.Value;
        var columnHeaderRowEnd = excelType.LastColumnHeaderRow()!.Value;

        foreach (var property in excelType.Properties)
        {
            var propHeaderRange = worksheet.Range(
                columnHeaderRowStart, property.ColumnNumber,
                columnHeaderRowEnd, property.ColumnNumber);

            propHeaderRange.Merge();
            propHeaderRange.Value = property.HeaderName;
            property.HeaderStyle?.Invoke(propHeaderRange.Style);
        }

        // Has no auto filter, skip
        if (!excelType.ColumnHeaderAutoFilter)
            return;

        var columnHeaderColumnStart = excelType.FirstColumnHeaderColumn()!.Value;
        var columnHeaderColumnEnd = excelType.LastColumnHeaderColumn()!.Value;

        var headerRange = worksheet.Range(
            columnHeaderRowStart, columnHeaderColumnStart,
            columnHeaderRowEnd, columnHeaderColumnEnd);

        headerRange.SetAutoFilter();
    }

    private static void WriteRow<T>(
        ExcelType excelType,
        (IXLWorksheet Worksheet, WorksheetMetadata Metadata) metadataMap,
        T entity,
        WorksheetOptions? options)
        where T : class
    {
        var nextEmptyRow = metadataMap.Worksheet.NextEmptyRow();

        var useOptions = options ?? WorksheetOptions.Default;
        foreach (var property in excelType.Properties)
        {
            var cell = nextEmptyRow.Cell(metadataMap.Metadata, property)
                       ?? throw new CellMismatchException(
                           new CellAddress(nextEmptyRow.RowNumber(), property.ColumnNumber), property.HeaderName);

            var value = property.PropertyInfo.GetValue(entity);

            if (value != null)
            {
                var ignore = property.Ignore?.Invoke(property.PropertyInfo, value, cell.Style, RowEvent.Write);
                if (ignore == true)
                    continue;
            }

            // Use custom serializer if exists
            XLCellValue cellValue;
            if (property.Serializer != null)
            {
                var context = new SerializationContext(useOptions, property.PropertyInfo, cell.Style);
                cellValue = property.Serializer.Invoke(value, context);
            }
            else
            {
                cellValue = value.ConvertToXLCellValue(property, useOptions);
            }

            cell.Value = cellValue;

            property.CellStyle?.Invoke(entity, property.PropertyInfo, value, new CellComment(cell), cell.Style);
        }
    }
}