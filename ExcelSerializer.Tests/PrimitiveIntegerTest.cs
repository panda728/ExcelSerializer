﻿// <auto-generated>
// THIS (.cs) FILE IS GENERATED BY `Serializers/PrimitiveSerializer.tt`. DO NOT CHANGE IT.
// </auto-generated>
#nullable enable
namespace ExcelSerializer.Tests
{
    public partial class PrimitiveSerializerTest
    {
        [Fact]
        public void Serializer_Byte()
        {
            var value1 = Byte.MinValue;
            var value2 = Byte.MaxValue;
            RunIntegerTest(value1, value2, ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_SByte()
        {
            var value1 = SByte.MinValue;
            var value2 = SByte.MaxValue;
            RunIntegerTest(value1, value2, ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_Int32()
        {
            var value1 = Int32.MinValue;
            var value2 = Int32.MaxValue;
            RunIntegerTest(value1, value2, ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_UInt32()
        {
            var value1 = UInt32.MinValue;
            var value2 = UInt32.MaxValue;
            RunIntegerTest(value1, value2, ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_Int64()
        {
            var value1 = Int64.MinValue;
            var value2 = Int64.MaxValue;
            RunIntegerTest(value1, value2, ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_UInt64()
        {
            var value1 = UInt64.MinValue;
            var value2 = UInt64.MaxValue;
            RunIntegerTest(value1, value2, ExcelSerializerOptions.Default);
        }
    }
}