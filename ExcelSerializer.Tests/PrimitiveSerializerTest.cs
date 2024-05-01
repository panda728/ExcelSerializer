using System.IO.Pipelines;
using System.Text;
using FluentAssertions;

namespace ExcelSerializerLib.Tests
{
    public partial class PrimitiveSerializerTest
    {
        internal static void RunIntegerTest<T>(T value1, T value2, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var ms = new MemoryStream();
            var writer = PipeWriter.Create(ms);
            var formatter = new ExcelFormatter(option);
            serializer.Serialize(ref formatter, writer, value1, option);
            serializer.Serialize(ref formatter, writer, value2, option);
            Assert.Empty(formatter.SharedStrings);
            writer.Complete();
            var result = Encoding.UTF8.GetString(ms.ToArray());
            result.Should().Be($"<c t=\"n\" s=\"5\"><v>{value1}</v></c><c t=\"n\" s=\"5\"><v>{value2}</v></c>");
        }
        internal static void RunNumberTest<T>(T value1, T value2, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var ms = new MemoryStream();
            var writer = PipeWriter.Create(ms);
            var formatter = new ExcelFormatter(option);
            serializer.Serialize(ref formatter, writer, value1, option);
            serializer.Serialize(ref formatter, writer, value2, option);
            Assert.Empty(formatter.SharedStrings);
            writer.Complete();
            var result = Encoding.UTF8.GetString(ms.ToArray());
            result.Should().Be($"<c t=\"n\" s=\"6\"><v>{value1}</v></c><c t=\"n\" s=\"6\"><v>{value2}</v></c>");
        }

        [Fact]
        public void Serializer_Boolean()
        {
            var option = ExcelSerializerOptions.Default;
            var serializer = option.GetSerializer<Boolean>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var value1 = true;
            var value2 = false;
            var ms = new MemoryStream();
            var writer = PipeWriter.Create(ms);
            var formatter = new ExcelFormatter(option);
            serializer.Serialize(ref formatter, writer, value1, option);
            serializer.Serialize(ref formatter, writer, value2, option);
            Assert.Empty(formatter.SharedStrings);
            writer.Complete();
            var result = Encoding.UTF8.GetString(ms.ToArray());
            result.Should().Be($"<c t=\"b\"><v>1</v></c><c t=\"b\"><v>0</v></c>");
        }
    }
}
