using System.Collections.Concurrent;

namespace ExcelSerializer.Providers;

public sealed class AdhocExcelSerializerProvider : IExcelSerializerProvider
{
    readonly IExcelSerializer[] serializers;
    readonly ConcurrentDictionary<Type, IExcelSerializer?> cache;
    readonly Func<Type, IExcelSerializer?> factory;

    public AdhocExcelSerializerProvider(IExcelSerializer[] serializers)
    {
        this.serializers = serializers;
        this.cache = new ConcurrentDictionary<Type, IExcelSerializer?>();
        this.factory = CreateSerializer;
    }

    public IExcelSerializer<T>? GetSerializer<T>() => (IExcelSerializer<T>?)cache.GetOrAdd(typeof(T), factory);

    IExcelSerializer? CreateSerializer(Type type)
    {
        foreach (var serializer in serializers)
        {
            var excelSerializerType = serializer.GetType().GetImplementedGenericType(typeof(IExcelSerializer<>));
            if (excelSerializerType != null && excelSerializerType.GenericTypeArguments[0] == type)
            {
                return serializer;
            }
        }

        return null;
    }
}
