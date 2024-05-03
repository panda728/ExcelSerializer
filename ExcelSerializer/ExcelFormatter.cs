using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelSerializerLib;

public sealed class ExcelFormatter
{
    //const int XF_NORMAL = 0;
    const int XF_WRAP_TEXT = 1;
    const int XF_DATETIME = 2;
    const int XF_DATE = 3;
    const int XF_INT = 5;
    const int XF_NUM = 6;

    const int LEN_DATE = 10;
    const int LEN_DATETIME = 18;

#if NET6_0_OR_GREATER
    const int XF_TIME = 4;
    const int LEN_TIME = 8;
#endif

    static readonly byte[] _emptyColumn = Encoding.UTF8.GetBytes("<c></c>");
    static readonly byte[] _colStartBoolean = Encoding.UTF8.GetBytes(@"<c t=""b""><v>");
    static readonly byte[] _colStartInteger = Encoding.UTF8.GetBytes(@$"<c t=""n"" s=""{XF_INT}""><v>");
    static readonly byte[] _colStartNumber = Encoding.UTF8.GetBytes(@$"<c t=""n"" s=""{XF_NUM}""><v>");
    static readonly byte[] _colStartStringWrap = Encoding.UTF8.GetBytes(@$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>");
    static readonly byte[] _colStartString = Encoding.UTF8.GetBytes(@$"<c t=""s""><v>");
    static readonly byte[] _colEnd = Encoding.UTF8.GetBytes(@"</v></c>");
    static readonly byte[] _boolTrue = Encoding.UTF8.GetBytes("1");
    static readonly byte[] _boolFalse = Encoding.UTF8.GetBytes("0");

    readonly ExcelSerializerOptions _options;

    bool _countingCharLength;

    int _columnIndex = 0;
    int _currentDepth = 0;
    int _stringIndex = 0;

    public ExcelFormatter(ExcelSerializerOptions options)
    {
        _options = options;
        _currentDepth = 0;
        _countingCharLength = options.AutoFitColumns;
    }

    /// <summary>
    /// Maintain a dictionary of strings. Output the same value with the same ID.
    /// </summary>
    public Dictionary<string, int> SharedStrings { get; } = new();

    /// <summary>
    /// Tally the maximum number of characters per column. For automatic column width adjustment
    /// </summary>
    public Dictionary<int, int> ColumnMaxLength { get; } = new();
    public void StopCountingCharLength() => _countingCharLength = false;

    public void Clear()
    {
        _columnIndex = 0;
        _currentDepth = 0;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void EnterAndValidate()
    {
        _currentDepth++;
        if (_currentDepth >= _options.MaxDepth)
            ThrowReachedMaxDepth(_currentDepth);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Exit()
    {
        _currentDepth--;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public int WriteEmpty(IBufferWriter<byte> writer)
    {
        _columnIndex++;
        writer.Write(_emptyColumn);
        return 0;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void SetMaxLength(int length)
    {
        if (!_countingCharLength)
            return;

#if NETSTANDARD2_1_OR_GREATER
        if (!ColumnMaxLength.TryAdd(_columnIndex, length))
        {
            if (ColumnMaxLength[_columnIndex] < length)
                ColumnMaxLength[_columnIndex] = length;
        }
#else
        if (ColumnMaxLength.TryGetValue(_columnIndex, out int value))
        {
            if (value < length)
                ColumnMaxLength[_columnIndex] = length;
        }
        else
        {
            ColumnMaxLength.Add(_columnIndex, length);
        }
#endif
        _columnIndex++;
    }

    //    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    //    void WriteRaw(string s, IBufferWriter<byte> writer)
    //    {
    //#if NET8_0_OR_GREATER
    //        var toReturn = ArrayPool<byte>.Shared.Rent(Encoding.UTF8.GetByteCount(s));
    //        var buffer = toReturn.AsSpan();
    //        Encoding.UTF8.TryGetBytes(s, toReturn, out var written);
    //        writer.Write(written, buffer[..written]);
    //        //writer.Advance(written);
    //        ArrayPool<byte>.Shared.Return(toReturn);
    //#elif NET5_0_OR_GREATER
    //        var bytes = Encoding.UTF8.GetBytes(s.AsSpan());
    //        writer.Write(bytes);
    //        //writer.Advance(bytes.Length);
    //#else
    //        var bytes = Encoding.UTF8.GetBytes(s);
    //        writer.Write(bytes);
    //        //writer.Advance(bytes.Length);
    //#endif
    //    }

    /// <summary>Write string.</summary>
    public void Write(string? s, IBufferWriter<byte> writer)
    {
        if (s == null || string.IsNullOrEmpty(s))
        {
            WriteEmpty(writer);
            return;
        }

        var index = SharedStrings.TryAdd(s, _stringIndex)
            ? _stringIndex++
            : SharedStrings[s];

        var span = s.Contains(Environment.NewLine) ? _colStartStringWrap.AsSpan() : _colStartString.AsSpan();
        writer.Write(span);
        WriteRaw($"{index}".AsSpan(), writer);
        writer.Write(_colEnd);
        SetMaxLength(s.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteRaw(ReadOnlySpan<char> chars, IBufferWriter<byte> writer)
    {
#if NET8_0_OR_GREATER
        var toReturn = ArrayPool<byte>.Shared.Rent(Encoding.UTF8.GetByteCount(chars));
        var buffer = toReturn.AsSpan();
        Encoding.UTF8.TryGetBytes(chars, toReturn, out var written);
        writer.Write(buffer[..written]);
        ArrayPool<byte>.Shared.Return(toReturn);
#else
        var bytes = Encoding.UTF8.GetBytes(chars.ToArray());
        writer.Write(bytes);
        //writer.Advance(bytes.Length);
#endif
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(char value, IBufferWriter<byte> writer) => Write($"{value}", writer);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(bool value, IBufferWriter<byte> writer)
    {
        writer.Write(_colStartBoolean);
        var span = value ? _boolTrue.AsSpan() : _boolFalse.AsSpan();
        writer.Write(span);
        writer.Write(_colEnd);
        SetMaxLength(1);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriterInteger(in ReadOnlySpan<char> chars, IBufferWriter<byte> writer)
    {
        writer.Write(_colStartInteger);
        WriteRaw(chars, writer);
        writer.Write(_colEnd);
        SetMaxLength(chars.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriterNumber(in ReadOnlySpan<char> chars, IBufferWriter<byte> writer)
    {
        writer.Write(_colStartNumber);
        WriteRaw(chars, writer);
        writer.Write(_colEnd);
        SetMaxLength(chars.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteDateTime(DateTime value, IBufferWriter<byte> writer)
    {
        var d = value;
        if (d == DateTime.MinValue) WriteEmpty(writer);
        if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
        {
            WriteRaw(@$"<c t=""d"" s=""{XF_DATE}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>".AsSpan(), writer);
            SetMaxLength(LEN_DATE);
            return;
        }

        WriteRaw(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>".AsSpan(), writer);
        SetMaxLength(LEN_DATETIME);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(byte value, IBufferWriter<byte> writer) => WriterInteger($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(sbyte value, IBufferWriter<byte> writer) => WriterInteger($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(decimal value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(double value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(float value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(int value, IBufferWriter<byte> writer) => WriterInteger($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(uint value, IBufferWriter<byte> writer) => WriterInteger($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(long value, IBufferWriter<byte> writer) => WriterInteger($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(ulong value, IBufferWriter<byte> writer) => WriterInteger($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(short value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(ushort value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);

#if NET5_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(Half value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);
#endif

#if NET6_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteDateTime(DateOnly value, IBufferWriter<byte> writer)
    {
        WriteRaw(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-dd}T00:00:00</v></c>".AsSpan(), writer);
        SetMaxLength(LEN_DATE);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteDateTime(TimeOnly value, IBufferWriter<byte> writer)
    {
        WriteRaw(@$"<c t=""d"" s=""{XF_TIME}""><v>1900-01-01T{value:HH:mm:ss}</v></c>".AsSpan(), writer);
        SetMaxLength(LEN_TIME);
    }
#endif

#if NET7_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(Int128 value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(UInt128 value, IBufferWriter<byte> writer) => WriterNumber($"{value}".AsSpan(), writer);
#endif

#if NETSTANDARD2_1_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowReachedMaxDepth(int depth)
    {
        throw new InvalidOperationException($"Serializer detects reached max depth:{depth}. Please check the circular reference.");
    }
}