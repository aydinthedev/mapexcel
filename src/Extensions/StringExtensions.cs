using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace MapExcel.Extensions;

internal static class StringExtensions
{
    internal static bool TryConvert<T>(this string value, [NotNullWhen(true)] out object? result)
        where T : struct => TryConvert<T>(value, CultureInfo.CurrentCulture, NumberStyles.Any, out result);

    internal static bool TryConvert<T>(this string value, CultureInfo culture, [NotNullWhen(true)] out object? result)
        where T : struct => TryConvert<T>(value, culture, NumberStyles.Any, out result);

    /// <summary>
    ///     Tries to convert string value to T type.
    ///     Supported types are:
    ///     Boolean, Char, SByte, Byte, Int16, UInt16, Int32, UInt32, Int64, UInt64, Single, Double, Decimal, DateTime.
    ///     Boolean conversion success if value is "TRUE", "FALSE", "1", "0" with StringComparison.OrdinalIgnoreCase.
    ///     DateTime conversion success only if value can be parsed with given culture.
    ///     Returns default value of T if string value is null or empty.
    ///     Returns null if conversion is not possible.
    /// </summary>
    internal static bool TryConvert<T>(
        this string value,
        CultureInfo culture,
        NumberStyles numberStyles,
        [NotNullWhen(true)] out object? result)
        where T : struct
    {
        if (string.IsNullOrEmpty(value))
        {
            result = default(T);
            return true;
        }

        result = typeof(T) switch
        {
            var type when type == typeof(bool) => TryConvertToBoolean(value, out var textToBoolean)
                ? textToBoolean
                : null,
            var type when type == typeof(char) => char.TryParse(value, out var textToChar)
                ? textToChar
                : null,
            var type when type == typeof(DateTime) =>
                DateTime.TryParse(value, culture, out var textToDateTime)
                    ? textToDateTime
                    : null,
            // Assume any other type is a number
            _ => double.TryParse(value, numberStyles, culture, out var textToDouble)
                ? textToDouble.TryCast<T>(out var doubleToType)
                    ? doubleToType
                    : null
                : null
        };

        return result != null;

        // Tries to convert string to a boolean value.
        //  Matches "TRUE", "FALSE", "1", "0" with StringComparison.OrdinalIgnoreCase.
        static bool TryConvertToBoolean(string value, out bool result)
        {
            if (value.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
                || value.Equals("1", StringComparison.OrdinalIgnoreCase))
            {
                result = true;
                return true;
            }

            if (value.Equals("FALSE", StringComparison.OrdinalIgnoreCase)
                || value.Equals("0", StringComparison.OrdinalIgnoreCase))
            {
                result = false;
                return true;
            }

            result = false;
            return false;
        }
    }
}