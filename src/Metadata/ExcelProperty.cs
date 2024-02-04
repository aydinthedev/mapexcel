using System.Reflection;
using ClosedXML.Excel;
using MapExcel.Constants;
using MapExcel.Extensions;

namespace MapExcel.Metadata;

internal sealed class ExcelProperty
{
    internal ExcelProperty(PropertyInfo propertyInfo)
    {
        ArgumentNullException.ThrowIfNull(propertyInfo);

        PropertyInfo = propertyInfo;
        ColumnNumber = WorksheetConstants.MinColumnNumber;
        HeaderMatch = HeaderMatch.Default;
        HeaderName = propertyInfo.Name;
        Type = propertyInfo.PropertyType;

        if (Type.TryGetUnderlyingType(out var underlyingPropertyType))
        {
            UnderlyingType = underlyingPropertyType;
            IsNullableType = true;
        }
        else
        {
            UnderlyingType = Type;
        }

        if (!Type.IsCollection())
            return;

        var collectionElementType = Type.GetCollectionElementType();
        CollectionElementType = collectionElementType.TryGetUnderlyingType(out var underlyingElementType)
            ? underlyingElementType
            : collectionElementType;

        IsCollectionType = true;
    }

    internal PropertyInfo PropertyInfo { get; }

    internal Type Type { get; }

    internal Type UnderlyingType { get; }

    internal Type? CollectionElementType { get; }

    internal bool IsNullableType { get; }

    internal bool IsCollectionType { get; }

    internal int ColumnNumber { get; set; }

    internal HeaderMatch HeaderMatch { get; set; }

    internal string HeaderName { get; set; }

    internal Action<IXLStyle>? HeaderStyle { get; set; }

    internal Func<PropertyInfo, object, IXLStyle, RowEvent, bool>? Ignore { get; set; }

    internal Func<object?, SerializationContext, XLCellValue>? Serializer { get; set; }

    internal Func<XLCellValue, DeserializationContext, object?>? Deserializer { get; set; }

    internal Action<object, PropertyInfo, object?, CellComment, IXLStyle>? CellStyle { get; set; }
}