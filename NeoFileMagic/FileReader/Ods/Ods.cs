namespace NeoFileMagic.FileReader.Ods;

using System;
using System.IO;
using System.IO.Compression;

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
}
