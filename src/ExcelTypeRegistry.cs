using MapExcel.Metadata;
using MapExcel.Metadata.Builders;

namespace MapExcel;

public static class ExcelTypeRegistry
{
    private static readonly Dictionary<Type, ExcelType> ExcelTypes = new();

    public static void Register<T>(Action<ExcelTypeBuilder<T>> configure) where T : class, new()
    {
        ArgumentNullException.ThrowIfNull(configure);

        var type = typeof(T);
        if (Contains(type))
            throw new ArgumentException($"Type '{type}' already registered.");

        var excelType = ExcelTypeInitializer.InitializeWithBuilder(configure);
        ExcelTypes.Add(type, excelType);
    }

    public static bool Contains<T>() => Contains(typeof(T));

    public static bool Contains(Type type)
    {
        ArgumentNullException.ThrowIfNull(type);
        return ExcelTypes.ContainsKey(type);
    }

    public static void Remove<T>() => Remove(typeof(T));

    public static void Remove(Type type)
    {
        if (!Contains(type)) return;
        ExcelTypes.Remove(type);
    }

    internal static ExcelType Get<T>() => Get(typeof(T));

    internal static ExcelType Get(Type type)
    {
        ArgumentNullException.ThrowIfNull(type);

        return ExcelTypes.TryGetValue(type, out var excelType)
            ? excelType
            : throw new ArgumentOutOfRangeException(nameof(type), $"Type '{type}' is not registered.");
    }
}