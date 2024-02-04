using ClosedXML.Excel;

namespace MapExcel.Metadata.Builders;

public sealed class HeaderBuilder
{
    private readonly ExcelProperty _excelProperty;

    internal HeaderBuilder(ExcelProperty excelProperty)
    {
        ArgumentNullException.ThrowIfNull(excelProperty);
        _excelProperty = excelProperty;
    }

    public HeaderBuilder Name(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Name cannot be empty or null.", nameof(name));

        _excelProperty.HeaderName = name.Trim();
        return this;
    }

    public HeaderBuilder Style(Action<IXLStyle> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        _excelProperty.HeaderStyle = configure;
        return this;
    }

    public HeaderBuilder Match(HeaderMatch match)
    {
        ArgumentNullException.ThrowIfNull(match);
        _excelProperty.HeaderMatch = match;
        return this;
    }
}