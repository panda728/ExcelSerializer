using System.Collections.ObjectModel;
using System.IO.Pipelines;
using System.Text;
using FluentAssertions;

namespace ExcelSerializerLib.Tests
{
    public class SerializersTest
    {
        static void RunTest<T>(T value, string value1ShouldBe, string columnXmlShouldBe, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var ms = new MemoryStream();
            var writer = PipeWriter.Create(ms);
            var formatter = new ExcelFormatter(option);
            try
            {
                serializer.Serialize(ref formatter, writer, value, option);
                Assert.Single(formatter.SharedStrings);
                writer.Complete();
                var result = Encoding.UTF8.GetString(ms.ToArray());
                var sharedString1 = formatter.SharedStrings.First().Key;

                result.Should().Be(columnXmlShouldBe);
                sharedString1.Should().Be(value1ShouldBe);
            }
            catch
            {
                throw;
            }
        }

        static void RunTest<T>(T value, string value1ShouldBe1, string value1ShouldBe2, string columnXmlShouldBe, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var ms = new MemoryStream();
            var writer = PipeWriter.Create(ms);
            var formatter = new ExcelFormatter(option);
            try
            {
                serializer.Serialize(ref formatter, writer, value, option);
                Assert.Equal(2, formatter.SharedStrings.Count);
                writer.Complete();
                var result = Encoding.UTF8.GetString(ms.ToArray());
                var sharedString1 = formatter.SharedStrings.First().Key;
                var sharedString2 = formatter.SharedStrings.Skip(1).First().Key;

                result.Should().Be(columnXmlShouldBe);
                sharedString1.Should().Be(value1ShouldBe1);
                sharedString2.Should().Be(value1ShouldBe2);
            }
            catch
            {
                throw;
            }
        }
        [Fact]
        public void Serializer_TCollection()
        {
            var dinosaurs = new Collection<string>
            {
                "Psitticosaurus",
                "Caudipteryx"
            };
            SerializersTest.RunTest(dinosaurs, "Psitticosaurus", "Caudipteryx",
                "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c>",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_IDictionary()
        {
            var dic = new Dictionary<string, int> { { "key1", 1 }, { "key2", 2 } };
            SerializersTest.RunTest(dic, "key1", "key2",
                "<c t=\"s\"><v>0</v></c><c t=\"n\" s=\"5\"><v>1</v></c><c t=\"s\"><v>1</v></c><c t=\"n\" s=\"5\"><v>2</v></c>",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_KeyValuePair()
        {
            var dic = new Dictionary<string, int> { { "key1", 1 }, { "key2", 2 } };
            SerializersTest.RunTest(dic.First(), "key1",
                "<c t=\"s\"><v>0</v></c><c t=\"n\" s=\"5\"><v>1</v></c>",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_ObjectFallback()
        {
            var value = (object)"key1";
            SerializersTest.RunTest(value, "key1", "<c t=\"s\"><v>0</v></c>", ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_CompiledObject()
        {
            var potals1 = new Portal { Name = "Portal1", Owner = null, Level = 8 };
            SerializersTest.CompiledObjectTest(potals1, "Portal1", "<c t=\"s\"><v>0</v></c><c></c><c t=\"n\" s=\"5\"><v>8</v></c>", ExcelSerializerOptions.Default);
        }

        static void CompiledObjectTest<T>(
            T value,
            string value1ShouldBe,
            string columnXmlShouldBe,
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
                serializer.Serialize(ref formatter, writer, value, option);
                Assert.Single(formatter.SharedStrings);

                writer.Complete();
                var result = Encoding.UTF8.GetString(ms.ToArray());
                var sharedString1 = formatter.SharedStrings.First().Key;

                result.Should().Be(columnXmlShouldBe);
                sharedString1.Should().Be(value1ShouldBe);
            }
            catch
            {
                throw;
            }
        }

        public class Portal
        {
            public string Name { get; set; } = "";
            public string? Owner { get; set; }
            public int Level { get; set; }
        }

    }
}
