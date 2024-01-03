using System.Reflection;

namespace MapExcel;

public class RowData<T> where T : class
{
    private readonly List<PropertyInfo> _invalidProperties;

    internal RowData(T entity, int rowNumber)
    {
        Entity = entity;
        RowNumber = rowNumber;
        _invalidProperties = new List<PropertyInfo>();
    }

    public T Entity { get; }

    public int RowNumber { get; }

    public IEnumerable<PropertyInfo> InvalidProperties => _invalidProperties;

    internal void InvalidateProperty(PropertyInfo propertyInfo) => _invalidProperties.Add(propertyInfo);
}