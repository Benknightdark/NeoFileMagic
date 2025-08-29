using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using NeoFileMagic.FileReader.Ods;
using NeoFileMagic.FileReader.Ods.Exception;
using Newtonsoft.Json;

public sealed class NeoOdsDeserializeTests
{
    private static MemoryStream BuildOds(string contentXml)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("content.xml");
            using var s = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            s.Write(contentXml);
        }
        ms.Position = 0;
        return ms;
    }

    private static string WrapContent(string innerTable)
    => $"""
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="S">{innerTable}</table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
""";

    public sealed class BasicRow
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public bool Active { get; set; }
        public double Score { get; set; }
        public DateTimeOffset When { get; set; }
        public TimeSpan Duration { get; set; }
    }

    [Fact]
    public void Deserialize_BasicMapping_Works()
    {
        var table = """
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Id</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>Name</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>Active</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>Score</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>When</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>Duration</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="float" office:value="1"/>
    <table:table-cell office:value-type="string"><text:p>Alice</text:p></table:table-cell>
    <table:table-cell office:value-type="boolean" office:boolean-value="true"/>
    <table:table-cell office:value-type="float" office:value="9.5"/>
    <table:table-cell office:value-type="date" office:date-value="2024-05-01T12:00:00Z"/>
    <table:table-cell office:value-type="time" office:time-value="PT1H2M3S"/>
  </table:table-row>
""";

        using var ods = BuildOds(WrapContent(table));
        var doc = NeoOds.Load(ods);
        var rows = NeoOds.DeserializeSheetOrThrow<BasicRow>(doc.Sheets[0]).ToList();

        Assert.Single(rows);
        var r = rows[0];
        Assert.Equal(1, r.Id);
        Assert.Equal("Alice", r.Name);
        Assert.True(r.Active);
        Assert.Equal(9.5, r.Score, 3);
        Assert.Equal(DateTimeOffset.Parse("2024-05-01T12:00:00Z"), r.When);
        Assert.Equal(TimeSpan.FromSeconds(3723), r.Duration);
    }

    public sealed class LocalizedRow
    {
        [JsonProperty(PropertyName = "代碼", Order = 0)]
        public string? Code { get; set; }

        [JsonProperty(PropertyName = "啟用", Order = 1)]
        public bool Enabled { get; set; }
    }

    [Fact]
    public void Deserialize_JsonPropertyNameMapping_Works()
    {
        var table = """
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>代碼</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>啟用</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>ST01</text:p></table:table-cell>
    <table:table-cell office:value-type="boolean" office:boolean-value="true"/>
  </table:table-row>
""";

        using var ods = BuildOds(WrapContent(table));
        var doc = NeoOds.Load(ods);
        var rows = NeoOds.DeserializeSheetOrThrow<LocalizedRow>(doc.Sheets[0]).ToList();

        Assert.Single(rows);
        Assert.Equal("ST01", rows[0].Code);
        Assert.True(rows[0].Enabled);
    }

    public sealed class OrderRow
    {
        [JsonProperty(PropertyName = "A", Order = 2)]
        public string? A { get; set; }
        [JsonProperty(PropertyName = "B", Order = 1)]
        public string? B { get; set; }
    }

    [Fact]
    public void Deserialize_HeaderOrderMismatch_Throws()
    {
        // 期望順序為 B, A（因 Order=1 在前），但實際表頭是 A, B → 應拋例外
        var table = """
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>B</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell>
  </table:table-row>
""";
        using var ods = BuildOds(WrapContent(table));
        var doc = NeoOds.Load(ods);
        Assert.Throws<OdsHeaderOrderMismatchException>(() =>
            NeoOds.DeserializeSheetOrThrow<OrderRow>(doc.Sheets[0]).ToList());
    }

    public sealed class MissingHeaderRow
    {
        public string? NeedMe { get; set; }
    }

    [Fact]
    public void Deserialize_MissingHeader_Throws()
    {
        var table = """
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Other</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell>
  </table:table-row>
""";
        using var ods = BuildOds(WrapContent(table));
        var doc = NeoOds.Load(ods);
        Assert.Throws<OdsHeaderMismatchException>(() =>
            NeoOds.DeserializeSheetOrThrow<MissingHeaderRow>(doc.Sheets[0]).ToList());
    }

    public sealed class StopRow
    {
        public string? Name { get; set; }
    }

    [Fact]
    public void Deserialize_StopAtFirstAllEmptyRow_Works()
    {
        var table = """
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Name</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>N1</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="string"></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>junk</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>N2</text:p></table:table-cell>
  </table:table-row>
""";

        using var ods = BuildOds(WrapContent(table));
        var doc = NeoOds.Load(ods);
        var rows = NeoOds.DeserializeSheetOrThrow<StopRow>(doc.Sheets[0], stopAtFirstAllEmptyRow: true).ToList();

        Assert.Single(rows);
        Assert.Equal("N1", rows[0].Name);
    }

    public sealed class ConvRow
    {
        public int N { get; set; }
    }

    [Fact]
    public void Deserialize_ConversionError_AggregatesAndThrows()
    {
        var table = """
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>N</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
    <table:table-cell office:value-type="string"><text:p>abc</text:p></table:table-cell>
  </table:table-row>
""";

        using var ods = BuildOds(WrapContent(table));
        var doc = NeoOds.Load(ods);
        var ex = Assert.Throws<OdsAggregateConversionException>(() =>
            NeoOds.DeserializeSheetOrThrow<ConvRow>(doc.Sheets[0]).ToList());

        Assert.True(ex.Errors.Count >= 1);
        var e = ex.Errors[0];
        Assert.Equal("N", e.JsonPropertyName);
        Assert.Equal(typeof(int), e.TargetType);
    }
}
