using ExcelSerializer;

namespace ExcelSerializerConsole;

//public class BoolZeroOneSerializer : IExcelSerializer<bool>
//{
//    public void WriteTitle(ref ExcelSerializerWriter writer, bool value, ExcelSerializerOptions options, string name = "")
//    {
//        writer.Write(name);
//    }

//    public void Serialize(ref ExcelSerializerWriter writer, bool value, ExcelSerializerOptions options)
//    {
//        // true => 0, false => 1
//        writer.WritePrimitive(value ? 0 : 1);
//    }
//}

public class UnixSecondsSerializer : IExcelSerializer<DateTime>
{
    public void WriteTitle(ref ExcelSerializerWriter writer, DateTime value, ExcelSerializerOptions options, string name = "")
    {
        writer.Write(name);
    }

    public void Serialize(ref ExcelSerializerWriter writer, DateTime value, ExcelSerializerOptions options)
    {
        writer.WritePrimitive(((DateTimeOffset)(value)).ToUnixTimeSeconds());
    }
}
