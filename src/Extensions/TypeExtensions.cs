using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace MapExcel.Extensions;

internal static class TypeExtensions
{
    internal static bool TryGetUnderlyingType(this Type type, [NotNullWhen(true)] out Type? underlyingType)
    {
        underlyingType = Nullable.GetUnderlyingType(type);
        return underlyingType != null;
    }

    internal static bool IsGenericCollection(this Type type) =>
        type.IsGenericType && type.IsCollection();

    internal static bool IsNonGenericCollection(this Type type) =>
        !type.IsGenericType && type.IsCollection();

    internal static bool IsCollection(this Type type) =>
        type.IsArray || (type != typeof(string) && typeof(IEnumerable).IsAssignableFrom(type));

    internal static Type GetCollectionElementType(this Type type)
    {
        if (type.IsArray)
            return type.GetElementType()!;

        // Return "object" type for non-generic collections like ArrayList
        return type.IsGenericCollection()
            ? type.GetGenericArguments()[0]
            : type.IsNonGenericCollection()
                ? typeof(object)
                : throw new NotSupportedException($"Type '{type.Name}' is not a supported collection type.");
    }

    internal static object CreateNewCollection(this Type type, int size)
    {
        if (type.IsArray)
        {
            var elementType = type.GetElementType()!;
            return Array.CreateInstance(elementType, size);
        }

        if (type.IsGenericCollection())
        {
            var genericType = type.GetGenericArguments()[0];
            var listType = typeof(List<>).MakeGenericType(genericType);
            return Activator.CreateInstance(listType, size)!;
        }

        if (type.IsNonGenericCollection())
            return Activator.CreateInstance(type, size)!;

        throw new NotSupportedException($"Type '{type.Name}' is not a supported collection type.");
    }
}