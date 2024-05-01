using System.Buffers;

namespace ExcelSerializerLib.Serializers;

internal class BuiltinSerializers
{
    public sealed class StringExcelSerializer : IExcelSerializer<string?>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, string? value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, string? value, ExcelSerializerOptions options)
            => formatter.Write(value, writer);
    }

    public sealed class CharExcelSerializer : IExcelSerializer<char>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, char value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, char value, ExcelSerializerOptions options)
            => formatter.Write(value, writer);
    }

    public sealed class GuidExcelSerializer : IExcelSerializer<Guid>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, Guid value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, Guid value, ExcelSerializerOptions options)
            => formatter.Write($"{value}", writer);
    }

    public sealed class EnumExcelSerializer : IExcelSerializer<Enum>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, Enum value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, Enum value, ExcelSerializerOptions options)
            => formatter.Write($"{value}", writer);
    }

    public sealed class DateTimeExcelSerializer : IExcelSerializer<DateTime>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, DateTime value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, DateTime value, ExcelSerializerOptions options)
            => formatter.WriteDateTime(value, writer);
    }

    public sealed class DateTimeOffsetExcelSerializer : IExcelSerializer<DateTimeOffset>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, DateTimeOffset value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, DateTimeOffset value, ExcelSerializerOptions options)
            => formatter.Write(value.ToString(options.CultureInfo), writer);
    }

    public sealed class TimeSpanExcelSerializer : IExcelSerializer<TimeSpan>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TimeSpan value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TimeSpan value, ExcelSerializerOptions options)
            => formatter.Write(value.ToString(), writer);
    }

    public sealed class UriExcelSerializer : IExcelSerializer<Uri>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, Uri value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, Uri value, ExcelSerializerOptions options)
        {
            if (value == null)
            {
                formatter.WriteEmpty(writer);
                return;
            }
            formatter.Write($"{value}", writer);
        }
    }

#if NET6_0_OR_GREATER
    public sealed class DateOnlyExcelSerializer : IExcelSerializer<DateOnly>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, DateOnly value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, DateOnly value, ExcelSerializerOptions options)
            => formatter.WriteDateTime(value, writer);
    }

    public sealed class TimeOnlyExcelSerializer : IExcelSerializer<TimeOnly>
    {
        public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TimeOnly value, ExcelSerializerOptions options, string name = "value")
            => formatter.Write(name, writer);
        public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, TimeOnly value, ExcelSerializerOptions options)
            => formatter.WriteDateTime(value, writer);
    }
#endif
}
