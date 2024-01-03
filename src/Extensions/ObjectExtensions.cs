using System.Collections;
using System.Text;
using ClosedXML.Excel;
using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class ObjectExtensions
{
    // TODO: Support enum and dictionaries
    internal static XLCellValue ConvertToXLCellValue(this object? obj, ExcelProperty property, WorksheetOptions options)
    {
        if (obj == null)
            return Blank.Value;

        if (!property.IsCollectionType)
            // Datetime values are dependent on the timezone, should be converted to the timezone of the worksheet
            return Type.GetTypeCode(property.UnderlyingType) switch
            {
                TypeCode.DateTime => ((DateTime)obj).AdjustToTimeZone(options.TimeZone),
                _ => XLCellValue.FromObject(obj)
            };

        var sb = new StringBuilder();
        foreach (var element in (IEnumerable)obj)
        {
            var elementValue = Type.GetTypeCode(property.CollectionElementType) switch
            {
                TypeCode.DateTime => ((DateTime)element).AdjustToTimeZone(options.TimeZone),
                _ => element
            };

            sb.Append(Convert.ToString(elementValue, options.Culture));
            sb.Append(options.CollectionSeparator);
        }

        if (sb.Length == 0)
            return Blank.Value;

        // Remove last separator
        sb.Length -= options.CollectionSeparator.Length;

        return sb.ToString();
    }
}