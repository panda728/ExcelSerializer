using System.Buffers;

namespace ExcelSerializerLib.Serializers;

public sealed class EnumerableExcelSerializer<TCollection, TElement> : IExcelSerializer<TCollection>
    where TCollection : IEnumerable<TElement>
{
    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TCollection value, ExcelSerializerOptions options, string name = "value")
    {
        formatter.EnterAndValidate();
        var serializer = options.GetRequiredSerializer<TElement>();
        foreach (var item in value)
        {
            serializer.WriteTitle(ref formatter, writer, item, options, name);
        }
        formatter.Exit();
    }

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TCollection value, ExcelSerializerOptions options)
    {
        if (value == null)
        {
            formatter.WriteEmpty(writer);
            return;
        }

        formatter.EnterAndValidate();
        var serializer = options.GetRequiredSerializer<TElement>();
        foreach (var item in value)
        {
            serializer.Serialize(ref formatter, writer, item, options);
        }
        formatter.Exit();
    }
}

public sealed class DictionaryExcelSerializer<TDictionary, TKey, TValue> : IExcelSerializer<TDictionary>
    where TDictionary : IDictionary<TKey, TValue>
{
    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TDictionary value, ExcelSerializerOptions options, string name = "value")
    {

        formatter.EnterAndValidate();
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(ref formatter, writer, item.Key, options, "key");
            valueSerializer.WriteTitle(ref formatter, writer, item.Value, options, name);
        }
        formatter.Exit();
    }

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TDictionary value, ExcelSerializerOptions options)
    {
        if (value == null)
        {
            formatter.WriteEmpty(writer);
            return;
        }

        formatter.EnterAndValidate();
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();

        foreach (var item in value)
        {
            if (item.Value == null)
            {
                formatter.WriteEmpty(writer);
                continue;
            }

            keySerializer.Serialize(ref formatter, writer, item.Key, options);
            valueSerializer.Serialize(ref formatter, writer, item.Value, options);
        }
        formatter.Exit();
    }
}

public sealed class EnumerableKeyValuePairExcelSerializer<TCollection, TKey, TValue> : IExcelSerializer<TCollection>
    where TCollection : IEnumerable<KeyValuePair<TKey, TValue>>
{
    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TCollection value, ExcelSerializerOptions options, string name = "value")
    {
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        formatter.EnterAndValidate();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(ref formatter, writer, item.Key, options, "key");
            valueSerializer.WriteTitle(ref formatter, writer, item.Value, options, name);
        }
        formatter.Exit();
    }

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TCollection value, ExcelSerializerOptions options)
    {
        if (value == null)
        {
            formatter.WriteEmpty(writer);
            return;
        }

        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        formatter.EnterAndValidate();
        foreach (var item in value)
        {
            if (item.Value == null)
            {
                formatter.WriteEmpty(writer);
                continue;
            }
            keySerializer.Serialize(ref formatter, writer, item.Key, options);
            valueSerializer.Serialize(ref formatter, writer, item.Value, options);
        }
        formatter.Exit();
    }
}