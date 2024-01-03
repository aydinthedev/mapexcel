using System.Diagnostics.CodeAnalysis;

namespace MapExcel.Extensions;

internal static class DoubleExtensions
{
    /// <summary>
    ///     Tries to cast double value to T type.
    ///     Supported types are:
    ///     Boolean, Char, SByte, Byte, Int16, UInt16, Int32, UInt32, Int64, UInt64, Single, Double, Decimal, DateTime.
    ///     Boolean cast is successful only if value is exact "1.0" or "0.0".
    ///     DateTime cast is successful only if value inbounds of OLE Automation Date range.
    ///     Returns null if cast is not possible.
    /// </summary>
    internal static bool TryCast<T>(this double value, [NotNullWhen(true)] out object? result) where T : struct
    {
        // If double is 1.0 or 0.0, cast it to boolean
        result = Type.GetTypeCode(typeof(T)) switch
        {
            // ReSharper disable once CompareOfFloatsByEqualityOperator
            TypeCode.Boolean when value is 1.0 or 0.0 => value == 1.0,
            TypeCode.Char when value is >= char.MinValue and <= char.MaxValue => (char)value,
            TypeCode.SByte when value is >= sbyte.MinValue and <= sbyte.MaxValue => (sbyte)value,
            TypeCode.Byte when value is >= byte.MinValue and <= byte.MaxValue => (byte)value,
            TypeCode.Int16 when value is >= short.MinValue and <= short.MaxValue => (short)value,
            TypeCode.UInt16 when value is >= ushort.MinValue and <= ushort.MaxValue => (ushort)value,
            TypeCode.Int32 when value is >= int.MinValue and <= int.MaxValue => (int)value,
            TypeCode.UInt32 when value is >= uint.MinValue and <= uint.MaxValue => (uint)value,
            TypeCode.Int64 when value is >= long.MinValue and <= long.MaxValue => (long)value,
            TypeCode.UInt64 when value is >= ulong.MinValue and <= ulong.MaxValue => (ulong)value,
            TypeCode.Single when value is >= float.MinValue and <= float.MaxValue => (float)value,
            TypeCode.Double => value,
            TypeCode.Decimal => (decimal)value,
            TypeCode.DateTime when value is >= -657434 and <= 2958465 => DateTime.FromOADate(value),
            _ => null
        };

        return result != null;
    }
}