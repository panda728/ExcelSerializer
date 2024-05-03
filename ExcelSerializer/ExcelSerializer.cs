using System.Buffers;
using System.IO.Compression;
using System.IO.Pipelines;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security;
using System.Text;

namespace ExcelSerializerLib;

public static class ExcelSerializer
{
    readonly static byte[] _contentTypes = Encoding.UTF8.GetBytes(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Override PartName=""/book.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/sheet.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
<Override PartName=""/strings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml""/>
<Override PartName=""/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");
    readonly static byte[] _rels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""book.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""/>
</Relationships>");

    readonly static byte[] _book = Encoding.UTF8.GetBytes(@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<bookViews><workbookView/></bookViews>
<sheets><sheet name=""Sheet"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>");
    readonly static byte[] _bookRels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""sheet.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""/>
<Relationship Id=""rId2"" Target=""strings.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings""/>
<Relationship Id=""rId3"" Target=""styles.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""/>
</Relationships>");

    readonly static string _styles = @"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<numFmts count=""5"">
<numFmt numFmtId=""1"" formatCode =""{0}"" />
<numFmt numFmtId=""2"" formatCode =""{1}"" />
<numFmt numFmtId=""3"" formatCode =""{2}"" />
<numFmt numFmtId=""4"" formatCode =""{3}"" />
<numFmt numFmtId=""5"" formatCode =""{4}"" />
</numFmts>
<fonts count=""1"">
<font/>
</fonts>
<fills count=""1"">
<fill/>
</fills>
<borders count=""1"">
<border/>
</borders>
<cellStyleXfs count=""1"">
<xf/>
</cellStyleXfs>
<cellXfs count=""7"">
<xf/>
<xf><alignment wrapText=""true""/></xf>
<xf numFmtId=""1""  applyNumberFormat=""1""></xf>
<xf numFmtId=""2""  applyNumberFormat=""1""></xf>
<xf numFmtId=""3""  applyNumberFormat=""1""></xf>
<xf numFmtId=""4""  applyNumberFormat=""1""></xf>
<xf numFmtId=""5""  applyNumberFormat=""1""></xf>
</cellXfs>
</styleSheet>";

    readonly static byte[] _rowStart = Encoding.UTF8.GetBytes("<row>");
    readonly static byte[] _rowEnd = Encoding.UTF8.GetBytes("</row>");
    readonly static byte[] _colStart = Encoding.UTF8.GetBytes("<cols>");
    readonly static byte[] _colEnd = Encoding.UTF8.GetBytes("</cols>");
    readonly static byte[] _frozenTitleRow = Encoding.UTF8.GetBytes(@"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>");

    readonly static byte[] _sheetStart = Encoding.UTF8.GetBytes(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">");
    readonly static byte[] _sheetEnd = Encoding.UTF8.GetBytes(@"</worksheet>");
    readonly static byte[] _dataStart = Encoding.UTF8.GetBytes(@"<sheetData>");
    readonly static byte[] _dataEnd = Encoding.UTF8.GetBytes(@"</sheetData>");

    readonly static byte[] _autoFilterStart = Encoding.UTF8.GetBytes(@"<autoFilter ref=""");
    readonly static byte[] _autoFilterEnd = Encoding.UTF8.GetBytes(@"""/>");

    readonly static byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
    //readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""1"">");
    readonly static byte[] _sstEnd = Encoding.UTF8.GetBytes(@"</sst>");
    readonly static byte[] _siStart = Encoding.UTF8.GetBytes("<si><t>");
    readonly static byte[] _siEnd = Encoding.UTF8.GetBytes("</t></si>");

    const int COLUMN_WIDTH_MARGIN = 2;
    private const string CONTENT_TYPE_XML = "[Content_Types].xml";
    private const string SHEET_XML = "sheet.xml";
    private const string STRINGS_XML = "strings.xml";
    private const string RELS = "_rels";
    private const string BOOK_XML = "book.xml";
    private const string STYLES_XML = "styles.xml";
    private const string BOOK_XML_RELS = "book.xml.rels";
    private const string DOT_RELS = ".rels";

    public static void ToFile<T>(IEnumerable<T> rows, string fileName, ExcelSerializerOptions options)
    {
        if (rows == null || !rows.Any())
            return;

        var workPathRoot = Path.Combine(options.WorkPath, Guid.NewGuid().ToString());
        if (!Directory.Exists(workPathRoot))
            Directory.CreateDirectory(workPathRoot);

        var formatter = new ExcelFormatter(options);
        try
        {
            using (var sheetStream = new FileStream(Path.Combine(workPathRoot, SHEET_XML), FileMode.Create, FileAccess.Write, FileShare.None))
            using (var stringsStream = new FileStream(Path.Combine(workPathRoot, STRINGS_XML), FileMode.Create, FileAccess.Write, FileShare.None))
            {
                var sheetWriter = PipeWriter.Create(sheetStream);
                CreateSheet(rows, formatter, sheetWriter, options);
                sheetWriter.Complete();

                var stringsWriter = PipeWriter.Create(stringsStream);
                WriteSharedStrings(formatter, stringsWriter);
                stringsWriter.Complete();
            }

            var workRelPath = Path.Combine(workPathRoot, RELS);
            if (!Directory.Exists(workRelPath))
                Directory.CreateDirectory(workRelPath);

            var _stylesBytes = Encoding.UTF8.GetBytes(string.Format(
                _styles,
                options.DateTimeFormat,
                options.DateFormat,
                options.TimeFormat,
                options.IntegerFormat,
                options.NumberFormat
            ));

            WriteStream(_contentTypes, Path.Combine(workPathRoot, CONTENT_TYPE_XML));
            WriteStream(_book, Path.Combine(workPathRoot, BOOK_XML));
            WriteStream(_stylesBytes, Path.Combine(workPathRoot, STYLES_XML));
            WriteStream(_rels, Path.Combine(workRelPath, DOT_RELS));
            WriteStream(_bookRels, Path.Combine(workRelPath, BOOK_XML_RELS));

            if (File.Exists(fileName))
                File.Delete(fileName);

            ZipFile.CreateFromDirectory(workPathRoot, fileName);
        }
        finally
        {
            try
            {
                if (Directory.Exists(workPathRoot))
                    Directory.Delete(workPathRoot, true);
            }
            catch { }
        }
    }

    static void WriteStream(byte[] bytes, string fileName)
    {
        using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
            fs.Write(bytes, 0, bytes.Length);
    }


    static void CreateSheet<T>(IEnumerable<T> rows, ExcelFormatter formatter, IBufferWriter<byte> writer, ExcelSerializerOptions options)
    {
        writer.Write(_sheetStart);

        if (options.HasHeaderRecord)
            writer.Write(_frozenTitleRow);

        if (options.AutoFitColumns)
            WriteCellWidth(rows, formatter, writer, options);

        writer.Write(_dataStart);
        //writer.Advance(_dataStart.Length);

        var serializer = options.GetSerializer<T>();
        if (serializer != null)
        {
            if (options.HasHeaderRecord)
            {
                writer.Write(_rowStart);
                //writer.Advance(_rowStart.Length);
                if (options.HeaderTitles != null && options.HeaderTitles.Length != 0)
                {
                    foreach (var t in options.HeaderTitles)
                        formatter.Write(t, writer);
                }
                else
                {
                    serializer.WriteTitle(ref formatter, writer, rows.First(), options);
                }
                writer.Write(_rowEnd);
                //writer.Advance(_rowEnd.Length);
            }

#if NET5_0_OR_GREATER
            if (rows is T[] arr)
                WriteRowsSpan(arr.AsSpan(), formatter, writer, serializer, options);
            else if (rows is List<T> list)
                WriteRowsSpan(CollectionsMarshal.AsSpan(list), formatter, writer, serializer, options);
            else
                WriteRows(rows, formatter, writer, serializer, options);
#else
            if (rows is T[] arr)
                WriteRowsSpan(arr.AsSpan(), formatter, writer, serializer, options);
            else
                WriteRows(rows, formatter, writer, serializer, options);
#endif
        }
        writer.Write(_dataEnd);

        if (options.AutoFilter)
        {
            var colName = options.HeaderTitles != null && options.HeaderTitles.Length != 0
                ? ToColumnName(options.HeaderTitles.Length)
                : ToColumnName(formatter.ColumnMaxLength.Count);

            var range = $"A1:{colName}{rows.Count() + 1}";
            writer.Write(_autoFilterStart);
            formatter.Write(range, writer);
            writer.Write(_autoFilterEnd);
        }
        writer.Write(_sheetEnd);
    }

    static string ToColumnName(int index)
    {
        if (index < 1) { return ""; }
        var list = new List<char>();
        index--;
        do
        {
            list.Add(Convert.ToChar(index % 26 + 65));
        }
        while ((index = index / 26 - 1) != -1);
        var sb = new StringBuilder(list.Count);
        for (int i = list.Count - 1; i >= 0; i--)
        {
            sb.Append(list[i]);
        }
        return sb.ToString();
    }

    static void WriteRowsSpan<T>(Span<T> rows, ExcelFormatter formatter, IBufferWriter<byte> writer, IExcelSerializer<T> serializer, ExcelSerializerOptions options)
    {
        foreach (var row in rows)
        {
            writer.Write(_rowStart);
            serializer.Serialize(ref formatter, writer, row, options);
            writer.Write(_rowEnd);
        }
    }

    static void WriteRows<T>(IEnumerable<T> rows, ExcelFormatter formatter, IBufferWriter<byte> writer, IExcelSerializer<T> serializer, ExcelSerializerOptions options)
    {
        foreach (var row in rows)
        {
            writer.Write(_rowStart);
            serializer.Serialize(ref formatter, writer, row, options);
            writer.Write(_rowEnd);
        }
    }

    static void WriteCellWidth<T>(IEnumerable<T> rows, ExcelFormatter formatter, IBufferWriter<byte> writer, ExcelSerializerOptions options)
    {
        // Counting the number of characters in Writer's internal process
        // The result is stored in writer.ColumnMaxLength 
        var serializer = options.GetSerializer<T>();
        if (serializer == null) return;
        if (options.HasHeaderRecord && options.HeaderTitles != null)
        {
            foreach (var t in options.HeaderTitles)
                formatter.Write(t, writer);
            formatter.Clear();
        }
        foreach (var row in rows.Take(options.AutoFitDepth))
        {
            serializer.Serialize(ref formatter, writer, row, options);
            formatter.Clear();
        }
        formatter.StopCountingCharLength();

        //var size = 100 * formatter.ColumnMaxLength.Count;
        writer.Write(_colStart);
        foreach (var pair in formatter.ColumnMaxLength)
        {
            var id = pair.Key + 1;
            var width = Math.Min(options.AutoFitWidhtMax, pair.Value + COLUMN_WIDTH_MARGIN);
            ExcelFormatter.WriteRaw(@$"<col min=""{id}"" max =""{id}"" width =""{width:0.0}"" bestFit =""1"" customWidth =""1"" />".AsSpan(), writer);
        }
        writer.Write(_colEnd);
    }

    static void WriteSharedStrings(ExcelFormatter formatter, IBufferWriter<byte> writer)
    {
        writer.Write(_sstStart);
        foreach (var s in formatter.SharedStrings.Keys)
        {
            writer.Write(_siStart);
            ExcelFormatter.WriteRaw(SecurityElement.Escape(s).AsSpan(), writer);
            writer.Write(_siEnd);
        }
        writer.Write(_sstEnd);
    }

    //    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    //    static void WriteRaw(string? s, ArrayPoolBufferWriter writer)
    //    {
    //        if (s == null)
    //            return;
    //#if NET5_0_OR_GREATER
    //        Encoding.UTF8.GetBytes(s.AsSpan(), writer);
    //#else
    //        writer.Write(Encoding.UTF8.GetBytes(s));
    //#endif
    //    }
}
