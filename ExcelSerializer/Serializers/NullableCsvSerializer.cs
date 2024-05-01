using System.Buffers;

namespace ExcelSerializerLib.Serializers;

public sealed class NullableExcelSerializer<T> : IExcelSerializer<T?>
    where T : struct
{
    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T? value, ExcelSerializerOptions options, string name = "value")
    {
        if (value == null)
        {
            formatter.Write(name, writer);
            return;
        }
        options.GetRequiredSerializer<T>().WriteTitle(ref formatter, writer, value.Value, options, name);
    }

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T? value, ExcelSerializerOptions options)
    {
        if (value == null)
        {
            formatter.WriteEmpty(writer);
            return;
        }
        options.GetRequiredSerializer<T>().Serialize(ref formatter, writer, value.Value, options);
    }
}