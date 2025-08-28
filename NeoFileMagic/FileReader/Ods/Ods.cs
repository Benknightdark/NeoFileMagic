namespace NeoFileMagic.FileReader.Ods;

using System;
using System.IO;
using System.IO.Compression;
using System.Net;

/// <summary>
/// 讀取 ODS 檔案的進入點。
/// </summary>
public static class Ods
{
    /// <summary>
    /// 從檔案路徑載入 ODS。
    /// </summary>
    public static OdsDocument Load(string path, OdsReaderOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path is null or empty", nameof(path));
        using var fs = File.OpenRead(path);
        return Load(fs, options);
    }

    /// <summary>
    /// 從串流載入 ODS（串流需可讀且可定位）。
    /// </summary>
    public static OdsDocument Load(Stream stream, OdsReaderOptions? options = null)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable", nameof(stream));
        options ??= new OdsReaderOptions();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // 安全性：偵測加密（manifest.xml 出現 manifest:encryption-data）
        var manifest = archive.GetEntry("META-INF/manifest.xml");
        if (manifest != null)
        {
            using var manStream = manifest.Open();
            if (OdsXml.HasEncryption(manStream))
            {
                if (options.ThrowOnEncrypted)
                    throw new NotSupportedException("Encrypted ODS is not supported.");
            }
        }

        var contentEntry = archive.GetEntry("content.xml") ?? throw new InvalidDataException("content.xml not found in ODS.");
        using var content = contentEntry.Open();
        var sheets = OdsXml.ParseContent(content, options);
        return new OdsDocument(sheets);
    }

    /// <summary>
    /// 以單行文字呈現儲存格內容，依指定策略處理換行與空白。
    /// </summary>
    /// <param name="cell">要輸出的儲存格。</param>
    /// <param name="mode">文字處理策略。</param>
    /// <returns>單行化後的字串。</returns>
    public static string OneLine(OdsCell cell, TextHandling mode = TextHandling.CollapseToSpace)
    {
        var s = cell.Type == OdsValueType.String ? (cell.Text ?? "") : cell.ToString();
        if (string.IsNullOrEmpty(s)) return "";
        s = s.Replace("\r\n", "\n").Replace('\r', '\n');

        return mode switch
        {
            TextHandling.Escape => s.Replace("\\", "\\\\").Replace("\t", "\\t").Replace("\n", "\\n"),
            TextHandling.FirstParagraph => (s.IndexOf('\n') is int idx && idx >= 0 ? s[..idx] : s).Trim(),
            TextHandling.CollapseToSpace =>
                string.Join(' ', s.Split(new[] { '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries))
                    .Replace("  ", " ").Trim(),
            _ => s
        };
    }

    // 建議共用 HttpClient 避免 socket 耗盡
    private static readonly HttpClient s_http = new(new HttpClientHandler
    {
        AutomaticDecompression = DecompressionMethods.All
    })
    {
        Timeout = TimeSpan.FromSeconds(100)
    };

    /// <summary>
    /// 從 URL 下載 ODS 並讀取（非同步）。
    /// - 僅支援 http/https/file
    /// - 內容長度大於 memoryThresholdBytes 或來源不可 Seek 時，會自動改寫入暫存檔避免佔用大量記憶體。
    /// </summary>
    public static async Task<OdsDocument> LoadFromUrlAsync(
        string url,
        OdsReaderOptions? options = null,
        CancellationToken cancellationToken = default,
        bool useTempFile = false,
        long memoryThresholdBytes = 64L * 1024 * 1024  // 64MB：超過則落地暫存檔
    )
    {
        if (string.IsNullOrWhiteSpace(url))
            throw new ArgumentException("URL is null or empty", nameof(url));

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            throw new ArgumentException("Invalid URL.", nameof(url));

        // 支援 file:// 直接走現有 Load
        if (uri.Scheme == Uri.UriSchemeFile)
            return Load(uri.LocalPath, options);

        if (uri.Scheme is not ("http" or "https"))
            throw new NotSupportedException("Only http/https/file schemes are supported.");

        using var req = new HttpRequestMessage(HttpMethod.Get, uri);
        req.Headers.UserAgent.ParseAdd("NeoFileMagic.OdsReader/1.0");

        using var resp = await s_http.SendAsync(req,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken).ConfigureAwait(false);

        resp.EnsureSuccessStatusCode();

        await using var netStream = await resp.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        var contentLength = resp.Content.Headers.ContentLength;

        // 需要 Seek 的 ZipArchive：若來源不可 Seek 或檔太大就用暫存檔
        if (useTempFile || !netStream.CanSeek || (contentLength.HasValue && contentLength.Value > memoryThresholdBytes))
        {
            var tempPath = Path.Combine(Path.GetTempPath(), $"neo_ods_{Guid.NewGuid():N}.ods");
            await using var fs = new FileStream(tempPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read,
                bufferSize: 81920, FileOptions.Asynchronous | FileOptions.DeleteOnClose);

            await netStream.CopyToAsync(fs, 81920, cancellationToken).ConfigureAwait(false);
            fs.Position = 0;
            // 直接用 Stream 版本（不需保留檔案）：Load 會在方法內把 ZIP 解析完畢
            return Load(fs, options);
        }
        else
        {
            // 已知長度則預配置容量，否則用預設
            using var ms = contentLength is > 0 && contentLength <= int.MaxValue
                ? new MemoryStream((int)contentLength.Value)
                : new MemoryStream();

            await netStream.CopyToAsync(ms, 81920, cancellationToken).ConfigureAwait(false);
            ms.Position = 0;
            return Load(ms, options);
    }
}


}
