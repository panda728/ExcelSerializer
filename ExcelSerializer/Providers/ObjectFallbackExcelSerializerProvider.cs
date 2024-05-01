using ExcelSerializerLib.Serializers;

namespace ExcelSerializerLib.Providers;

public class ObjectFallbackExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new ObjectFallbackExcelSerializerProvider();

    ObjectFallbackExcelSerializerProvider()
    {

    }

    public IExcelSerializer<T>? GetSerializer<T>() 
        => typeof(T) == typeof(object) ? (IExcelSerializer<T>)new ObjectFallbackExcelSerializer() : null;
}
