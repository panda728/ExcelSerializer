namespace ExcelSerializer;

internal static class ReflectionExtensions
{
    public static bool IsNullable(this Type type) 
        => Nullable.GetUnderlyingType(type) != null;

    public static Type? GetImplementedGenericType(this Type type, Type genericTypeDefinition) 
        => type.GetInterfaces().FirstOrDefault(x => x.IsConstructedGenericType && x.GetGenericTypeDefinition() == genericTypeDefinition);
}