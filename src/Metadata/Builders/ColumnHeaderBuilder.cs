using ClosedXML.Excel;

namespace MapExcel.Metadata.Builders;

public sealed class ColumnHeaderBuilder
{
    private readonly ExcelProperty _excelProperty;

    internal ColumnHeaderBuilder(ExcelProperty excelProperty)
    {
        ArgumentNullException.ThrowIfNull(excelProperty);
        _excelProperty = excelProperty;
    }

    public ColumnHeaderBuilder Name(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Name cannot be empty or null.", nameof(name));

        _excelProperty.HeaderName = name.Trim();
        return this;
    }

    public ColumnHeaderBuilder Style(Action<IXLStyle> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        _excelProperty.HeaderStyle = configure;
        return this;
    }
}