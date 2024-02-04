using System.Reflection;
using ClosedXML.Excel;
using MapExcel.Constants;

namespace MapExcel.Metadata.Builders;

public sealed class ExcelPropertyBuilder<T, TP> where T : class
{
    private readonly ExcelProperty _excelProperty;

    internal ExcelPropertyBuilder(ExcelProperty excelProperty)
    {
        ArgumentNullException.ThrowIfNull(excelProperty);
        _excelProperty = excelProperty;
    }

    public ExcelPropertyBuilder<T, TP> Column(int columnNumber)
    {
        if (columnNumber is < WorksheetConstants.MinColumnNumber or > WorksheetConstants.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnNumber));

        _excelProperty.ColumnNumber = columnNumber;
        return this;
    }

    public ExcelPropertyBuilder<T, TP> Column(string columnName)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty or null.", nameof(columnName));

        var columnNumber = CellAddress.ToColumnNumber(columnName);
        if (columnNumber is < WorksheetConstants.MinColumnNumber or > WorksheetConstants.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnName));

        _excelProperty.ColumnNumber = columnNumber;
        return this;
    }
    
    public ExcelPropertyBuilder<T, TP> Header(Action<HeaderBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        configure(new HeaderBuilder(_excelProperty));
        return this;
    }

    public ExcelPropertyBuilder<T, TP> CellStyle(Action<T, PropertyInfo, TP?, CellComment, IXLStyle> style)
    {
        ArgumentNullException.ThrowIfNull(style);
        _excelProperty.CellStyle = (o, p, v, c, s) => style((T)o, p, (TP?)v, c, s);
        return this;
    }

    public ExcelPropertyBuilder<T, TP> Ignore(Func<PropertyInfo, object, IXLStyle, RowEvent, bool> ignore)
    {
        ArgumentNullException.ThrowIfNull(ignore);
        _excelProperty.Ignore = ignore;
        return this;
    }

    public ExcelPropertyBuilder<T, TP> Ignore(Func<TP, IXLStyle, RowEvent, bool> ignore)
    {
        ArgumentNullException.ThrowIfNull(ignore);
        _excelProperty.Ignore = (_, val, style, ev) => ignore((TP)val, style, ev);
        return this;
    }

    public ExcelPropertyBuilder<T, TP> Serializer(Func<TP?, SerializationContext, XLCellValue> serialize)
    {
        ArgumentNullException.ThrowIfNull(serialize);
        _excelProperty.Serializer = (val, context) => serialize((TP?)val, context);
        return this;
    }

    public ExcelPropertyBuilder<T, TP> Deserializer(Func<XLCellValue, DeserializationContext, TP?>? deserialize)
    {
        ArgumentNullException.ThrowIfNull(deserialize);
        _excelProperty.Deserializer = (cellVal, context) => deserialize(cellVal, context);
        return this;
    }
}