using ClosedXML.Excel;
using MapExcel.Extensions;
using MapExcel.Metadata;

namespace MapExcel;

public sealed class ExcelWorkbook : IDisposable
{
    private readonly IXLWorkbook _workbook;
    private readonly WorksheetMetadataMap _worksheetMetadataMap;

    public ExcelWorkbook()
    {
        _workbook = new XLWorkbook();
        _worksheetMetadataMap = new WorksheetMetadataMap(_workbook);
    }

    public ExcelWorkbook(string filePath)
    {
        _workbook = new XLWorkbook(filePath);
        _worksheetMetadataMap = new WorksheetMetadataMap(_workbook);
    }

    public ExcelWorkbook(Stream fileStream)
    {
        _workbook = new XLWorkbook(fileStream);
        _worksheetMetadataMap = new WorksheetMetadataMap(_workbook);
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }

    public WorksheetMetadata? GetMetadata<T>() where T : class =>
        _worksheetMetadataMap.TryGet(ExcelTypeRegistry.Get(typeof(T)));

    public IEnumerable<RowData<T>> Read<T>(WorksheetOptions? options = null) where T : class
    {
        var typeMap = ExcelTypeRegistry.Get(typeof(T));

        var metadata =
            _worksheetMetadataMap.TryGet(typeMap)
            ?? throw new InvalidOperationException("Associated worksheet is not found");

        if (!metadata.Validate())
            throw new Exception("Worksheet is not valid");

        var firstDataRowFound = metadata.Worksheet.FirstDataRowFound(metadata);
        if (firstDataRowFound == null)
            yield break;

        var useOptions = options ?? WorksheetOptions.Default;
        foreach (var row in metadata.Worksheet.RowsUsed(x => x.RowNumber() >= firstDataRowFound.RowNumber()))
        {
            var entity = Activator.CreateInstance<T>();
            var rowData = new RowData<T>(entity, row.RowNumber());

            foreach (var property in typeMap.Properties)
            {
                var cell = row.Cell(metadata, property);

                // Use custom deserializer if exists
                object? value;
                if (property.Deserializer != null)
                {
                    var context = new DeserializationContext(useOptions, property.PropertyInfo, cell.Style);
                    value = property.Deserializer.Invoke(cell.Value, context);

                    // User can set context.IsValid
                    // If its set to true then conversion is failed so we must add error to row data
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
        var excelType = ExcelTypeRegistry.Get<T>();
        Write(excelType, options, entity);
    }

    public void WriteBulk<T>(IEnumerable<T> entities, WorksheetOptions? options = null) where T : class
    {
        var excelType = ExcelTypeRegistry.Get<T>();
        Write(excelType, options, entities.ToArray());
    }

    public void Save() => _workbook.Save();

    public void SaveAs(string filePath) => _workbook.SaveAs(filePath);

    public void SaveAs(Stream stream) => _workbook.SaveAs(stream);

    private void Write<T>(ExcelType excelType, WorksheetOptions? options = null, params T[] entities) where T : class
    {
        var metadata = _worksheetMetadataMap.GetOrAdd(excelType);

        if (!metadata.Validate())
            throw new Exception("Worksheet is not valid");

        foreach (var entity in entities)
            WriteRow(excelType, metadata, entity, options);
    }

    private static void WriteRow<T>(
        ExcelType excelType,
        WorksheetMetadata metadata,
        T entity,
        WorksheetOptions? options)
        where T : class
    {
        ArgumentNullException.ThrowIfNull(entity);

        // Ensure template is written to the worksheet
        WriteTemplate(metadata);

        var nextEmptyRow = metadata.Worksheet.NextEmptyRow();

        var useOptions = options ?? WorksheetOptions.Default;
        foreach (var property in excelType.Properties)
        {
            var cell = nextEmptyRow.Cell(metadata, property);
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

    private static void WriteTemplate(WorksheetMetadata metadata)
    {
        // Do not override if worksheet is not empty
        if (!metadata.Worksheet.IsEmpty())
            return;

        foreach (var caption in metadata.ExcelType.Captions)
        {
            var range = metadata.Worksheet.Range(caption.AddressRange);

            range.Merge();
            range.SetValue(caption.Name);
            caption.Style?.Invoke(range.Style);
        }

        if (!metadata.ExcelType.HasHeaders())
        {
            metadata.Update();
            return;
        }

        // All headers are in same row
        // But we need to calculate used column range for headers
        var headersExpectedAt = metadata.ExpectedHeaders[0].ExpectedAt!.Value;
        var firstHeaderRowNumber = headersExpectedAt.FirstCellAddress.RowNumber;
        var firstHeaderColumnNumber = headersExpectedAt.FirstCellAddress.ColumnNumber;
        var lastHeaderRowNumber = headersExpectedAt.LastCellAddress.RowNumber;
        var lastHeaderColumnNumber = headersExpectedAt.LastCellAddress.ColumnNumber;

        foreach (var property in metadata.ExcelType.Properties)
        {
            var range = metadata.Worksheet.Range(
                firstHeaderRowNumber,
                property.ColumnNumber,
                lastHeaderRowNumber,
                property.ColumnNumber);

            range.Merge();
            range.Value = property.HeaderName;
            property.HeaderStyle?.Invoke(range.Style);

            // Find used column range for headers
            if (property.ColumnNumber < firstHeaderColumnNumber)
                firstHeaderColumnNumber = property.ColumnNumber;

            if (property.ColumnNumber > lastHeaderColumnNumber)
                lastHeaderColumnNumber = property.ColumnNumber;
        }

        // Has no auto filter, skip
        if (!metadata.ExcelType.HeaderAutoFilter)
        {
            metadata.Update();
            return;
        }

        var headerRange = metadata.Worksheet.Range(
            firstHeaderRowNumber,
            firstHeaderColumnNumber,
            lastHeaderRowNumber,
            lastHeaderColumnNumber);

        headerRange.SetAutoFilter();

        metadata.Update();
    }
}