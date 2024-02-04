using System.Reflection;

namespace MapExcel.Metadata;

public class HeaderMetadata
{
    // Prevent outside initialization
    internal HeaderMetadata(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        Name = name;
    }

    internal ExcelProperty? ExcelProperty { get; init; }

    public string Name { get; }

    public PropertyInfo? Property => ExcelProperty?.PropertyInfo;

    public CellAddressRange? ExpectedAt { get; init; }

    public CellAddressRange? FoundAt { get; internal set; }

    public bool Validate()
    {
        if (FoundAt == null || ExpectedAt == null || ExcelProperty == null)
            return false;

        return ExcelProperty.HeaderMatch switch
        {
            // When ColumnMatch is HeaderOnly, we only check if the header name is found.
            HeaderMatch.NameOnly => true,

            // When ColumnMatch is HeaderAndColumn, we check if the header name found at expected position.
            HeaderMatch.NameAndColumn =>
                ExpectedAt.Value.FirstCellAddress.ColumnNumber == FoundAt.Value.FirstCellAddress.ColumnNumber,

            _ => throw new ArgumentOutOfRangeException()
        };
    }
}