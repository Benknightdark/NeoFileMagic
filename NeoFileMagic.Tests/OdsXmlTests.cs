using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using NeoFileMagic.FileReader.Ods;

/// <summary>
/// ODS 讀取與行為的核心測試。
/// 涵蓋基本型別解析、單行文字處理、加密偵測、
/// 檔案載入（暫存與 repo 範例檔）以及外部資料集 URL 檢查與可選下載。
/// </summary>
public sealed class OdsXmlTests
{
    /// <summary>
    /// 驗證 content.xml 解析基本型別：字串（含 text:s 連續空白、段落換行）、
    /// 浮點數、布林、日期（ISO-8601）與時間（TimeSpan），並測試 OneLine 各模式輸出。
    /// </summary>
    [Fact]
    public void OdsXml_ParseContent_BasicTypes_ReturnsValues()
    {
        var content = """
        <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="Sheet1">
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>Hello<text:s text:c="2"/>World</text:p><text:p>Second</text:p></table:table-cell>
                  <table:table-cell office:value-type="float" office:value="3.14"/>
                  <table:table-cell office:value-type="boolean" office:boolean-value="true"/>
                  <table:table-cell office:value-type="date" office:date-value="2024-01-02T03:04:05Z"/>
                  <table:table-cell office:value-type="time" office:time-value="PT1H2M3S"/>
                </table:table-row>
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
        """;

        using var ods = BuildOds(content);
        var doc = Ods.Load(ods);

        Assert.Single(doc.Sheets);
        var sh = doc.Sheets[0];
        Assert.Equal("Sheet1", sh.Name);
        Assert.Equal(1, sh.RowCount);

        var c0 = sh.GetCell(0, 0);
        Assert.Equal(OdsValueType.String, c0.Type);
        Assert.Equal("Hello  World\nSecond", c0.Text);
        Assert.Equal("Hello  World\\nSecond", Ods.OneLine(c0, TextHandling.Escape));
        Assert.Equal("Hello World Second", Ods.OneLine(c0, TextHandling.CollapseToSpace));
        Assert.Equal("Hello  World", Ods.OneLine(c0, TextHandling.FirstParagraph));

        var c1 = sh.GetCell(0, 1);
        Assert.Equal(OdsValueType.Float, c1.Type);
        Assert.Equal("3.14", c1.ToString());

        var c2 = sh.GetCell(0, 2);
        Assert.Equal(OdsValueType.Boolean, c2.Type);
        Assert.Equal("True", c2.ToString());

        var c3 = sh.GetCell(0, 3);
        Assert.Equal(OdsValueType.Date, c3.Type);
        Assert.Equal("2024-01-02T03:04:05.0000000+00:00", c3.ToString());

        var c4 = sh.GetCell(0, 4);
        Assert.Equal(OdsValueType.Time, c4.Type);
        Assert.Equal(TimeSpan.Parse("01:02:03").ToString(), c4.ToString());
    }

    /// <summary>
    /// 以記憶體組裝最小 ODS，寫到暫存路徑後用檔案 API 載入，
    /// 確認基本檔案路徑讀取流程（Zip + content.xml）。
    /// </summary>
    [Fact]
    public void SampleOds_File_Loads()
    {
        var content = """
        <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
          <office:body>
            <office:spreadsheet>
              <table:table table:name="SheetA">
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>Hi</text:p></table:table-cell>
                </table:table-row>
              </table:table>
            </office:spreadsheet>
          </office:body>
        </office:document-content>
        """;

        using var odsStream = BuildOds(content);
        var tmp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".ods");
        using (var fs = File.Create(tmp)) odsStream.CopyTo(fs);

        var doc = Ods.Load(tmp);
        Assert.Single(doc.Sheets);
        Assert.Equal("SheetA", doc.Sheets[0].Name);
        Assert.Equal("Hi", doc.Sheets[0].GetCell(0, 0).Text);
    }

    /// <summary>
    /// 若輸出目錄存在 repo 內提供的 sample.ods，則嘗試載入並確認結構；
    /// 若不存在（如 CI 環境），則不當作失敗以避免對環境產生耦合。
    /// </summary>
    [Fact]
    public void SampleOds_FromRepo_Loads_IfPresent()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "sample.ods");
        if (!File.Exists(path)) return; // CI 或開發環境未附帶檔案時，直接返回不失敗
        var doc = Ods.Load(path);
        Assert.NotNull(doc);
        Assert.NotNull(doc.Sheets);
    }

    /// <summary>
    /// 驗證資料集連結檔存在且內容為有效 URL，並包含指定 rId。
    /// 僅檢查檔案與字串格式，不進行網路下載。
    /// </summary>
    [Fact]
    public void Dataset_Url_File_Exists_And_Valid()
    {
        var urlFile = Path.Combine(AppContext.BaseDirectory, "Data", "DatasetUrl.txt");
        Assert.True(File.Exists(urlFile));
        var url = File.ReadAllText(urlFile).Trim();
        Assert.StartsWith("http", url, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("A21030000I-D20003-004", url);
    }

    [Fact(Skip = "需要網路才可下載驗證，預設略過。設定環境變數 NFM_ALLOW_NET=1 後改為手動啟用。")]
    public async System.Threading.Tasks.Task Dataset_Url_Download_Smoke()
    {
        if (Environment.GetEnvironmentVariable("NFM_ALLOW_NET") != "1")
            return; // 安全防護：即使移除 Skip 也不下載

        var urlFile = Path.Combine(AppContext.BaseDirectory, "Data", "DatasetUrl.txt");
        var url = File.ReadAllText(urlFile).Trim();
        using var http = new System.Net.Http.HttpClient();
        var resp = await http.GetAsync(url);
        resp.EnsureSuccessStatusCode();
        var json = await resp.Content.ReadAsStringAsync();
        Assert.False(string.IsNullOrWhiteSpace(json));
    }

    

    /// <summary>
    /// 檢驗 manifest 中的加密標記會在預設下拋出 NotSupportedException；
    /// 若配置 ThrowOnEncrypted=false，則允許載入但不輸出任何工作表。
    /// </summary>
    [Fact]
    public void Ods_Load_Encrypted_ThrowsByDefault_AllowsWhenDisabled()
    {
        var content = """
        <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <office:body><office:spreadsheet/></office:body>
        </office:document-content>
        """;

        var manifest = """
        <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml">
            <manifest:encryption-data/>
          </manifest:file-entry>
        </manifest:manifest>
        """;

        using var ods = BuildOds(content, manifest);
        Assert.Throws<NotSupportedException>(() => Ods.Load(ods));

        ods.Position = 0;
        var doc = Ods.Load(ods, new OdsReaderOptions { ThrowOnEncrypted = false });
        Assert.Empty(doc.Sheets);
    }

    /// <summary>
    /// 驗證單行文字輸出模式的行為：Escape、CollapseToSpace、FirstParagraph。
    /// </summary>
    [Fact]
    public void Ods_OneLine_Modes_Work()
    {
        var cell = new OdsCell { Type = OdsValueType.String, Text = "A\nB\tC\\D" };
        Assert.Equal("A", Ods.OneLine(new OdsCell { Type = OdsValueType.String, Text = "A\r\n" }, TextHandling.CollapseToSpace));
        Assert.Equal("A\\nB\\tC\\\\D", Ods.OneLine(cell, TextHandling.Escape));
        Assert.Equal("A B C\\D", Ods.OneLine(cell, TextHandling.CollapseToSpace));
        Assert.Equal("A", Ods.OneLine(cell, TextHandling.FirstParagraph));
    }

    /// <summary>
    /// 以最小結構在記憶體組裝 ODS：建立 Zip 並寫入 content.xml，
    /// 可選擇加入 META-INF/manifest.xml。
    /// </summary>
    /// <param name="contentXml">content.xml 內容字串。</param>
    /// <param name="manifestXml">manifest.xml 內容字串（可為 null 表示不加入）。</param>
    /// <returns>可供 Ods.Load 使用的 ODS Zip 記憶體串流。</returns>
    private static MemoryStream BuildOds(string contentXml, string? manifestXml = null)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("content.xml");
            using (var s = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
                s.Write(contentXml);

            if (manifestXml != null)
            {
                var man = zip.CreateEntry("META-INF/manifest.xml");
                using var sm = new StreamWriter(man.Open(), new UTF8Encoding(false));
                sm.Write(manifestXml);
            }
        }
        ms.Position = 0;
        return ms;
    }
}
