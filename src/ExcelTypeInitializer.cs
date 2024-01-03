using MapExcel.Metadata;
using MapExcel.Metadata.Builders;

namespace MapExcel;

internal static class ExcelTypeInitializer
{
    internal static ExcelType InitializeWithBuilder<T>(Action<ExcelTypeBuilder<T>> configure) where T : class, new()
    {
        var excelType = new ExcelType(typeof(T));
        var builder = new ExcelTypeBuilder<T>(excelType);

        configure(builder);

        return excelType;
    }

    // TODO: Annotation support
    internal static ExcelType InitializeWithAnnotations() => throw new NotImplementedException();
}