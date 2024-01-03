using System.Globalization;
using ClosedXML.Excel;
using MapExcel.Metadata;

namespace MapExcel.Extensions;

internal static class XLCellExtensions
{
    // TODO: Support enum and dictionaries
    internal static bool TryConvertToObject(
        this IXLCell cell,
        ExcelProperty excelProperty,
        WorksheetOptions options,
        out object? result)
    {
        if (!excelProperty.IsCollectionType)
        {
            // Nullable primitive excelProperty should be null if cell is empty
            if (excelProperty.IsNullableType && cell.Value.IsBlank)
            {
                result = null;
                return true;
            }

            // Convert the cell value to ExcelProperty type by remaining faithful to WorksheetOptions.
            // Each convert method has its own logic to do conversion that relies on cell value and number format
            switch (Type.GetTypeCode(excelProperty.UnderlyingType))
            {
                case TypeCode.String:
                    return cell.TryConvertToString(options.ConvertStringFrom, out result);
                case TypeCode.DateTime:
                    return cell.TryConvertToDateTime(
                        options.Culture, options.TimeZone, options.DateTimeKind, options.ConvertDateTimeFrom,
                        out result);
                case TypeCode.Boolean:
                    return cell.TryConvertToBoolean(options.ConvertBooleanFrom, out result);
                case TypeCode.Char:
                    return cell.TryConvertToChar(options.ConvertCharFrom, out result);
                case TypeCode.SByte:
                    return cell.TryConvertToNumber<sbyte>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Byte:
                    return cell.TryConvertToNumber<byte>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Int16:
                    return cell.TryConvertToNumber<short>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.UInt16:
                    return cell.TryConvertToNumber<ushort>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Int32:
                    return cell.TryConvertToNumber<int>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.UInt32:
                    return cell.TryConvertToNumber<uint>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Int64:
                    return cell.TryConvertToNumber<long>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.UInt64:
                    return cell.TryConvertToNumber<ulong>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Single:
                    return cell.TryConvertToNumber<float>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Double:
                    return cell.TryConvertToNumber<double>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Decimal:
                    return cell.TryConvertToNumber<decimal>(
                        options.Culture, options.NumberStyles, options.ConvertNumberFrom, out result);
                case TypeCode.Empty:
                case TypeCode.Object:
                case TypeCode.DBNull:
                default:
                    result = null;
                    return false;
            }
        }

        // Split cell value to collection of strings
        var values = cell.GetString().Split(
            options.CollectionSeparator,
            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

        var collection = excelProperty.Type.CreateNewCollection(values.Length);

        // If cell is empty then return empty collection
        if (values.Length < 1)
        {
            result = collection;
            return true;
        }

        Action<object, int> collectionAdd = excelProperty.Type.IsArray
            ? (item, index) =>
                (excelProperty.Type.GetMethod(nameof(Array.SetValue), new[] { typeof(object), typeof(int) }) ??
                 throw new MissingMethodException(excelProperty.Type.Name, nameof(Array.SetValue)))
                .Invoke(collection, new[] { item, index })
            : (item, _) =>
                (collection.GetType().GetMethod(nameof(List<object>.Add)) ??
                 throw new MissingMethodException(excelProperty.Type.Name, nameof(List<object>.Add)))
                .Invoke(collection, new[] { item });

        for (var i = 0; i < values.Length; i++)
        {
            object? value;
            switch (Type.GetTypeCode(excelProperty.CollectionElementType))
            {
                case TypeCode.Object:
                case TypeCode.String:
                    value = values[i];
                    break;
                case TypeCode.DateTime:
                    if (values[i].TryConvert<DateTime>(options.Culture, out var dateTime))
                        value = ((DateTime)dateTime).AdjustToKind(options.TimeZone, options.DateTimeKind);
                    else
                        value = null;
                    break;
                case TypeCode.Boolean:
                    values[i].TryConvert<bool>(out value);
                    break;
                case TypeCode.Char:
                    values[i].TryConvert<char>(out value);
                    break;
                case TypeCode.SByte:
                    values[i].TryConvert<sbyte>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Byte:
                    values[i].TryConvert<byte>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Int16:
                    values[i].TryConvert<short>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.UInt16:
                    values[i].TryConvert<ushort>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Int32:
                    values[i].TryConvert<int>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.UInt32:
                    values[i].TryConvert<uint>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Int64:
                    values[i].TryConvert<long>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.UInt64:
                    values[i].TryConvert<ulong>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Single:
                    values[i].TryConvert<float>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Double:
                    values[i].TryConvert<double>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Decimal:
                    values[i].TryConvert<decimal>(options.Culture, options.NumberStyles, out value);
                    break;
                case TypeCode.Empty:
                case TypeCode.DBNull:
                default:
                    value = null;
                    break;
            }

            if (value == null)
            {
                result = null;
                return false;
            }

            collectionAdd(value, i);
        }

        result = collection;
        return true;
    }

    internal static bool TryConvertToString(this IXLCell cell, StringFrom convertStringFrom, out object? result)
    {
        switch (cell.DataType)
        {
            case XLDataType.Blank:
                result = string.Empty;
                return true;
            case XLDataType.Text:
                result = cell.GetString().Trim();
                return true;

            // Cell is not blank or text, so convert only if non-text formatted cells are allowed
            case XLDataType.Boolean:
            case XLDataType.Number:
            case XLDataType.Error:
            case XLDataType.DateTime:
            case XLDataType.TimeSpan:
            default:
                if (convertStringFrom.HasFlag(StringFrom.NonText))
                {
                    result = cell.GetString().Trim();
                    return true;
                }

                result = null;
                return false;
        }
    }

    internal static bool TryConvertToBoolean(this IXLCell cell, BooleanFrom convertBooleanFrom, out object? result)
    {
        switch (cell.DataType)
        {
            case XLDataType.Blank:
                result = default(bool);
                return true;
            case XLDataType.Boolean:
                result = cell.GetBoolean();
                return true;
            case XLDataType.Number:
                if (convertBooleanFrom.HasFlag(BooleanFrom.Number))
                    return cell.GetDouble().TryCast<bool>(out result);
                result = null;
                return false;
            case XLDataType.Text:
                if (convertBooleanFrom.HasFlag(BooleanFrom.Text))
                    return cell.GetString().Trim().TryConvert<bool>(out result);
                result = null;
                return false;
            case XLDataType.Error:
            case XLDataType.DateTime:
            case XLDataType.TimeSpan:
            default:
                result = null;
                return false;
        }
    }

    internal static bool TryConvertToChar(this IXLCell cell, CharFrom convertCharFrom, out object? result)
    {
        switch (cell.DataType)
        {
            case XLDataType.Blank:
                result = default(char);
                return true;
            case XLDataType.Number:
                if (convertCharFrom.HasFlag(CharFrom.Number))
                    return cell.GetDouble().TryCast<char>(out result);
                result = null;
                return false;
            case XLDataType.Text:
                if (convertCharFrom.HasFlag(CharFrom.Text))
                    return cell.GetString().Trim().TryConvert<char>(out result);
                result = null;
                return false;
            case XLDataType.Boolean:
            case XLDataType.Error:
            case XLDataType.DateTime:
            case XLDataType.TimeSpan:
            default:
                result = null;
                return false;
        }
    }

    internal static bool TryConvertToDateTime(
        this IXLCell cell,
        CultureInfo culture,
        TimeZoneInfo timeZone,
        DateTimeKind kind,
        DateTimeFrom convertDateTimeFrom,
        out object? result)
    {
        switch (cell.DataType)
        {
            case XLDataType.Blank:
                result = default(DateTime).AdjustToKind(timeZone, kind);
                return true;
            case XLDataType.DateTime:
                result = cell.GetDateTime().AdjustToKind(timeZone, kind);
                return true;
            case XLDataType.Number:
                if (convertDateTimeFrom.HasFlag(DateTimeFrom.Number) &&
                    cell.GetDouble().TryCast<DateTime>(out var dateFromNumber))
                {
                    result = ((DateTime)dateFromNumber).AdjustToKind(timeZone, kind);
                    return true;
                }

                result = null;
                return false;
            case XLDataType.Text:
                if (convertDateTimeFrom.HasFlag(DateTimeFrom.Text) &&
                    cell.GetString().Trim().TryConvert<DateTime>(culture, out var dateFromText))
                {
                    result = ((DateTime)dateFromText).AdjustToKind(timeZone, kind);
                    return true;
                }

                result = null;
                return false;
            case XLDataType.Boolean:
            case XLDataType.Error:
            case XLDataType.TimeSpan:
            default:
                result = null;
                return false;
        }
    }

    internal static bool TryConvertToNumber<T>(
        this IXLCell cell,
        CultureInfo culture,
        NumberStyles numberStyles,
        NumberFrom convertNumberFrom,
        out object? result)
        where T : struct
    {
        switch (cell.DataType)
        {
            case XLDataType.Blank:
                result = default(T);
                return true;
            case XLDataType.Number:
                return cell.GetDouble().TryCast<T>(out result);
            case XLDataType.TimeSpan:
                if (convertNumberFrom.HasFlag(NumberFrom.TimeSpan))
                    return cell.Value.GetUnifiedNumber().TryCast<T>(out result);
                result = null;
                return false;
            case XLDataType.DateTime:
                if (convertNumberFrom.HasFlag(NumberFrom.DateTime))
                    return cell.Value.GetUnifiedNumber().TryCast<T>(out result);
                result = null;
                return false;
            case XLDataType.Text:
                if (convertNumberFrom.HasFlag(NumberFrom.Text))
                    return cell.GetString().Trim().TryConvert<T>(culture, numberStyles, out result);
                result = null;
                return false;
            case XLDataType.Boolean:
            case XLDataType.Error:
            default:
                result = null;
                return false;
        }
    }
}