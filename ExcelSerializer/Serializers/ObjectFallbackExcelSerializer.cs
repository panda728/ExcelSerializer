using System.Buffers;
using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace ExcelSerializerLib.Serializers;

internal class ObjectFallbackExcelSerializer : IExcelSerializer<object>
{
    private delegate void WriteTitleDelegate(ref ExcelFormatter formatter, IBufferWriter<byte> writer, object value, ExcelSerializerOptions options, string name);
    static readonly ConcurrentDictionary<Type, WriteTitleDelegate> nongenericWriteTitles = new();
    static readonly Func<Type, WriteTitleDelegate> factoryWriteTitle = CompileWriteTitleDelegate;

    private delegate void SerializeDelegate(ref ExcelFormatter formatter, IBufferWriter<byte> writer, object value, ExcelSerializerOptions options);
    static readonly ConcurrentDictionary<Type, SerializeDelegate> nongenericSerializers = new();
    static readonly Func<Type, SerializeDelegate> factory = CompileSerializeDelegate;

    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, object value, ExcelSerializerOptions options, string name = "value")
    {
        var type = value.GetType();
        if (type == typeof(object))
        {
            formatter.Write(name, writer);
            return;
        }

        var writeTitle = nongenericWriteTitles.GetOrAdd(type, factoryWriteTitle);
        writeTitle.Invoke(ref formatter, writer, value, options, name);
    }

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, object value, ExcelSerializerOptions options)
    {
        if (value == null)
        {
            formatter.WriteEmpty(writer);
            return;
        }

        var type = value.GetType();
        if (type == typeof(object))
        {
            formatter.WriteEmpty(writer);
            return;
        }

        var serializer = nongenericSerializers.GetOrAdd(type, factory);
        serializer.Invoke(ref formatter, writer, value, options);
    }

    static WriteTitleDelegate CompileWriteTitleDelegate(Type type)
    {
        var formatter = Expression.Parameter(typeof(ExcelFormatter).MakeByRefType());
        var writer = Expression.Parameter(typeof(IBufferWriter<byte>));
        var value = Expression.Parameter(typeof(object));
        var options = Expression.Parameter(typeof(ExcelSerializerOptions));
        var name = Expression.Parameter(typeof(string));

        var getRequiredSerializer = typeof(ExcelSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        var writeTitle = typeof(IExcelSerializer<>).MakeGenericType(type).GetMethod("WriteTitle")!;
        var body = Expression.Call(
            Expression.Call(options, getRequiredSerializer),
            writeTitle,
            formatter,
            writer,
            Expression.Convert(value, type),
            options,
            name);

        var lambda = Expression.Lambda<WriteTitleDelegate>(body, formatter, writer, value, options, name);
        return lambda.Compile();
    }

    static SerializeDelegate CompileSerializeDelegate(Type type)
    {
        // Serialize(ref ExcelSerializerWriter writer, object value, ExcelSerializerOptions options)
        //   options.GetRequiredSerializer<T>().Serialize(ref writer, (T)value, options)

        var formatter = Expression.Parameter(typeof(ExcelFormatter).MakeByRefType());
        var writer = Expression.Parameter(typeof(IBufferWriter<byte>));
        var value = Expression.Parameter(typeof(object));
        var options = Expression.Parameter(typeof(ExcelSerializerOptions));

        var getRequiredSerializer = typeof(ExcelSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        var serialize = typeof(IExcelSerializer<>).MakeGenericType(type).GetMethod("Serialize")!;

        var body = Expression.Call(
            Expression.Call(options, getRequiredSerializer),
            serialize,
            formatter,
            writer,
            Expression.Convert(value, type),
            options);

        var lambda = Expression.Lambda<SerializeDelegate>(body, formatter, writer, value, options);
        return lambda.Compile();
    }
}