using System.Buffers;
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.Serialization;

namespace ExcelSerializerLib.Serializers;

public sealed class EnumStringExcelSerializer<T> : IExcelSerializer<T>
    where T : Enum
{
    static readonly ConcurrentDictionary<T, string> stringCache = new();
    static readonly Func<T, string> toStringFactory = EnumToString;

    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options, string name = "value")
        => formatter.Write(name, writer);

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options)
    {
        var str = stringCache.GetOrAdd(value, toStringFactory);
        formatter.Write(str, writer);
    }

    static string EnumToString(T value)
    {
        var str = value.ToString();
        var field = value.GetType().GetField(str);
        if (field != null)
        {
            var enumMember = field.GetCustomAttribute<EnumMemberAttribute>();
            if (enumMember != null && enumMember.Value != null)
            {
                str = enumMember.Value;
            }
        }
        return str;
    }
}

public sealed class EnumValueExcelSerializer<T> : IExcelSerializer<T>
    where T : Enum
{
    static readonly ConcurrentDictionary<T, string> stringCache = new();
    static readonly Func<T, string> toStringFactory = EnumToString;

    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options, string name = "value")
        => formatter.Write(name, writer);

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options)
    {
        var str = stringCache.GetOrAdd(value, toStringFactory);
        formatter.Write(str, writer);
    }

    static string EnumToString(T value)
    {
        return Convert.ChangeType(value, Enum.GetUnderlyingType(typeof(T))).ToString()!;
    }
}