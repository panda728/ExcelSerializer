using FluentAssertions;
using ExcelSerializerLib;
using System.IO.Pipelines;
using System.Text;

namespace ExcelSerializerLib.Tests;

public class BuiltinSerializersTest
{
    static void RunStringColumnTest<T>(
        T value1, T value2,
        string value1ShouldBe, string value2ShouldBe,
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
            serializer.Serialize(ref formatter,writer, value1, option);
            serializer.Serialize(ref formatter, writer, value2, option);
            serializer.Serialize(ref formatter, writer, value1, option);

            Assert.Equal(2, formatter.SharedStrings.Count);

            writer.Complete();
            var columnXml = Encoding.UTF8.GetString(ms.ToArray());

            var sharedString1 = formatter.SharedStrings.First().Key;
            var sharedString2 = formatter.SharedStrings.Skip(1).First().Key;

            columnXml.Should().Be("<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>0</v></c>");
            sharedString1.Should().Be(value1ShouldBe);
            sharedString2.Should().Be(value2ShouldBe);
        }
        catch
        {
            throw;
        }
    }

    static void RunColumnTest<T>(
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
            writer.Complete();
            var result = Encoding.UTF8.GetString(ms.ToArray());

            Assert.Empty(formatter.SharedStrings);

            result.Should().Be(value1ShouldBe);
        }
        catch
        {
            throw;
        }
    }

    [Fact]
    public void Serializer_string()
    {
        BuiltinSerializersTest.RunStringColumnTest(
            "column1", "column2",
            "column1", "column2",
            ExcelSerializerOptions.Default);
    }

    [Fact]
    public void Serializer_char()
    {
        BuiltinSerializersTest.RunStringColumnTest(
            'A', 'Z',
            "A", "Z",
            ExcelSerializerOptions.Default);
    }

    [Fact]
    public void Serializer_Guid()
    {
        var guid1 = Guid.NewGuid();
        var guid2 = Guid.NewGuid();

        BuiltinSerializersTest.RunStringColumnTest(
            guid1, guid2,
            guid1.ToString(), guid2.ToString(),
            ExcelSerializerOptions.Default);
    }

    private enum DayOfWeek
    {
        Mon, Tue, Wed, Thu, Fri, Sat, Sun
    }

    [Fact]
    public void Serializer_Enum()
    {
        BuiltinSerializersTest.RunStringColumnTest(
            DayOfWeek.Mon, DayOfWeek.Tue,
            DayOfWeek.Mon.ToString(), DayOfWeek.Tue.ToString(),
            ExcelSerializerOptions.Default);
    }

    [Fact]
    public void Serializer_DateTime()
    {
        var value = new DateTime(2000, 1, 1);
        BuiltinSerializersTest.RunColumnTest(
            value,
            "<c t=\"d\" s=\"3\"><v>2000-01-01T00:00:00</v></c>",
            ExcelSerializerOptions.Default);
    }

    [Fact]
    public void Serializer_DateTimeOffset()
    {
        var option = ExcelSerializerOptions.Default;
        var value1 = DateTimeOffset.Now;
        var value2 = DateTimeOffset.UtcNow;
        BuiltinSerializersTest.RunStringColumnTest(
            value1, value2,
            value1.ToString(option.CultureInfo), value2.ToString(option.CultureInfo),
            option);
    }

    [Fact]
    public void Serializer_TimeSpan()
    {
        var value1 = DateTime.Today.AddHours(10) - DateTime.Today;
        var value2 = DateTime.Today.AddHours(-10) - DateTime.Today;
        BuiltinSerializersTest.RunStringColumnTest(
            value1, value2,
            "10:00:00", "-10:00:00",
            ExcelSerializerOptions.Default);
    }
    [Fact]
    public void Serializer_Uri()
    {
        var value1 = new Uri("http://hoge.com/fuga");
        var value2 = new Uri("http://hoge.com/fugafuga");
        BuiltinSerializersTest.RunStringColumnTest(
            value1, value2,
            "http://hoge.com/fuga", "http://hoge.com/fugafuga",
            ExcelSerializerOptions.Default);
    }
#if NET6_0_OR_GREATER
    [Fact]
    public void Serializer_DateOnly()
    {
        var option = ExcelSerializerOptions.Default;
        var value1 = DateOnly.FromDateTime(new DateTime(2000, 1, 1));
        var value2 = DateOnly.FromDateTime(new DateTime(9999, 12, 31));
        BuiltinSerializersTest.RunColumnTest(value1, "<c t=\"d\" s=\"3\"><v>2000-01-01T00:00:00</v></c>", option);
        BuiltinSerializersTest.RunColumnTest(value2, "<c t=\"d\" s=\"3\"><v>9999-12-31T00:00:00</v></c>", option);
    }
    [Fact]
    public void Serializer_TimeOnly()
    {
        var option = ExcelSerializerOptions.Default;
        var value1 = TimeOnly.FromDateTime(new DateTime(2000, 1, 1, 0, 0, 0));
        var value2 = TimeOnly.FromDateTime(new DateTime(9999, 12, 31, 23, 59, 59));
        BuiltinSerializersTest.RunColumnTest(value1, "<c t=\"d\" s=\"4\"><v>1900-01-01T00:00:00</v></c>", option);
        BuiltinSerializersTest.RunColumnTest(value2, "<c t=\"d\" s=\"4\"><v>1900-01-01T23:59:59</v></c>", option);
    }
#endif
}
