using System.Linq.Expressions;
using System.Reflection;

namespace MapExcel.Metadata.Builders;

public sealed class ExcelTypeBuilder<T> where T : class, new()
{
    private readonly ExcelType _excelType;

    internal ExcelTypeBuilder(ExcelType excelType)
    {
        ArgumentNullException.ThrowIfNull(excelType);
        _excelType = excelType;
    }

    public ExcelTypeBuilder<T> Worksheet(Action<WorksheetBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        configure(new WorksheetBuilder(_excelType));
        return this;
    }

    public ExcelTypeBuilder<T> Headers(int rowCount)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));

        _excelType.ColumnHeaderCount = rowCount;
        return this;
    }

    public ExcelTypeBuilder<T> AutoFilter()
    {
        _excelType.ColumnHeaderAutoFilter = true;
        return this;
    }

    public ExcelTypeBuilder<T> Property<TP>(
        Expression<Func<T, TP>> property,
        Action<ExcelPropertyBuilder<T, TP>> configure)
    {
        ArgumentNullException.ThrowIfNull(property);
        ArgumentNullException.ThrowIfNull(configure);

        if (property.Body is not MemberExpression memberExpression)
            throw new ArgumentException("Expression must be a member expression", nameof(property));

        if (memberExpression.Member is not PropertyInfo propertyInfo)
            throw new ArgumentException("Expression must represent a property", nameof(property));

        if (_excelType.Properties.Any(p => p.PropertyInfo.Name == propertyInfo.Name))
            throw new InvalidOperationException(
                $"Property '{propertyInfo.Name}' already registered for type '{_excelType.Type}'.");

        var excelProperty = new ExcelProperty(propertyInfo);
        var builder = new ExcelPropertyBuilder<T, TP>(excelProperty);
        configure(builder);

        _excelType.Properties.Add(excelProperty);
        return this;
    }
}