using System.IO.Pipelines;
using System.Text;
using FluentAssertions;
using Xunit.Sdk;

namespace ExcelSerializerLib.Tests
{
    public partial class TupleSerializersTest
    {
        static void RunTest<T>(
            T value1, string value1ShouldBe,
            ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var ms = new MemoryStream();
            var writer = PipeWriter.Create(ms);
            var formatter = new ExcelFormatter(option);
            try
            {
                serializer.Serialize(ref formatter, writer, value1, option);
                Assert.Empty(formatter.SharedStrings);
                writer.Complete();
                var result = Encoding.UTF8.GetString(ms.ToArray());
                result.Should().Be(value1ShouldBe);
            }
            catch
            {
                throw;
            }
        }
    }
}
