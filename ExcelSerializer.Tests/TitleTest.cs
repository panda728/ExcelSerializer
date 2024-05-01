using System.IO.Pipelines;
using System.Runtime.Serialization;
using System.Text;
using FluentAssertions;

namespace ExcelSerializerLib.Tests;

public class TitleTest
{
    static void RunStringColumnTest<T>(
        T value1,
        string value1ShouldBe,
        string value2ShouldBe,
        string value3ShouldBe,
        string titleShouldBe,
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
            serializer.WriteTitle(ref formatter ,writer, value1, option);
            Assert.Equal(3, formatter.SharedStrings.Count);

            writer.Complete();
            var result = Encoding.UTF8.GetString(ms.ToArray());
            var sharedString1 = formatter.SharedStrings.First().Key;
            var sharedString2 = formatter.SharedStrings.Skip(1).First().Key;
            var sharedString3 = formatter.SharedStrings.Skip(2).First().Key;

            result.Should().Be(titleShouldBe);
            sharedString1.Should().Be(value1ShouldBe);
            sharedString2.Should().Be(value2ShouldBe);
            sharedString3.Should().Be(value3ShouldBe);
        }
        catch
        {
            throw;
        }
    }

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
            serializer.WriteTitle(ref formatter, writer, value, option);
            Assert.NotEmpty(formatter.SharedStrings);
            var sharedString1 = formatter.SharedStrings.First().Key;
            writer.Complete();
            var result = Encoding.UTF8.GetString(ms.ToArray());

            result.Should().Be(columnXmlShouldBe);
            sharedString1.Should().Be(value1ShouldBe);
        }
        catch
        {
            throw;
        }
    }

    [Fact]
    public void Serializer_WriteTitle()
    {
        var list = new List<TestData>()
        {
            new(){Title  = "Title1", Name = "Name1", Address="Address1"},
            new(){Title  = "Title2", Name = "Name2", Address="Address2"},
            new(){Title  = "Title3", Name = "Name3", Address="Address3"},
        };

        var option = ExcelSerializerOptions.Default with
        {
            HasHeaderRecord = true,
        };

        TitleTest.RunStringColumnTest(
            list,
            "Address Ex",
            "Title Ex",
            "Name Ex",
            "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>2</v></c><c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>2</v></c><c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>2</v></c>",
            option);
    }

    [Fact]
    public void Serializer_ObjectFallback()
    {
        var value = (object)"key1";
        TitleTest.RunTest(value, "value",
            "<c t=\"s\"><v>0</v></c>",
            ExcelSerializerOptions.Default);
    }

    [Fact]
    public void Serializer_tuple2()
    {
        var t = Tuple.Create(1, 2);
        TitleTest.RunTest(t, "value", "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c>", ExcelSerializerOptions.Default);
    }
    [Fact]
    public void Serializer_IDictionary()
    {
        var dic = new Dictionary<string, int> { { "key1", 1 }, { "key2", 2 } };
        TitleTest.RunTest(dic, "key",
            "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c>",
            ExcelSerializerOptions.Default);
    }
}

public class TestData
{
    [DataMember(Name = "Title Ex", Order = 2)]
    public string Title { get; set; } = "";
    [DataMember(Name = "Name Ex", Order = 3)]
    public string Name { get; set; } = "";
    [DataMember(Name = "Address Ex", Order = 1)]
    public string Address { get; set; } = "";
}
