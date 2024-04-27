namespace ExcelSerializer.Providers;

public sealed partial class PrimitiveExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new PrimitiveExcelSerializerProvider();
    readonly Dictionary<Type, IExcelSerializer> serializers = new();

    internal partial void InitPrimitives(); // implement from PrimitiveSerializers.cs

    PrimitiveExcelSerializerProvider()
    {
        InitPrimitives();
    }

    public IExcelSerializer<T>? GetSerializer<T>() 
        => serializers.TryGetValue(typeof(T), out var value) ? (IExcelSerializer<T>)value : null;
}