using System.Buffers;

namespace ExcelSerializerLib;

public interface IExcelSerializer { }

public interface IExcelSerializer<T> : IExcelSerializer
{
    void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options, string name = "value");
    void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options);
}
