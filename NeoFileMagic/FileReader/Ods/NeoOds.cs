namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using NeoFileMagic.FileReader.Ods.Exception;
using Newtonsoft.Json;

/// <summary>
/// 讀取 ODS 檔案的進入點。
/// </summary>
public static class NeoOds
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



    /// <summary>
    /// 嚴格依表頭與 [JsonPropertyName] 對應，將工作表轉型為 T；
    /// 表頭缺漏、欄位順序不同或任一列資料格式不符時，拋出具體例外。
    /// </summary>
    /// <param name="sheet">來源工作表。</param>
    /// <param name="headerRowIndex">表頭列索引（預設 0）。</param>
    /// <param name="dataStartRowIndex">資料起始列索引（預設 1）。</param>
    /// <param name="caseInsensitiveHeaderMatch">表頭比對是否大小寫不敏感（預設 true）。</param>
    /// <param name="headerNormalizer">表頭正規化（預設 Trim）。</param>
    /// <param name="cellString">文字正規化函式（預設：\r\n/\r → \n，換行/Tab 摺疊為單一空白，最後 Trim）。</param>
    /// <param name="stopAtFirstAllEmptyRow">遇到第一個全空列即停止（預設 true）。</param>
    /// <param name="enforceHeaderOrder">
    /// 是否檢查欄位順序（預設 true）。
    /// 規則：以模型屬性順序（支援 [JsonPropertyOrder]；否則以宣告順序近似）作為期望順序，
    /// 實際表頭中對應之子序列若與期望順序不同則拋 
    /// <see cref="OdsHeaderOrderMismatchException"/>。
    /// </param>
    public static IEnumerable<T> DeserializeSheetOrThrow<T>(
        OdsSheet sheet,
        int headerRowIndex = 0,
        int dataStartRowIndex = 1,
        bool caseInsensitiveHeaderMatch = true,
        Func<string, string>? headerNormalizer = null,
        Func<OdsCell, string>? cellString = null,
        bool stopAtFirstAllEmptyRow = true,
        bool enforceHeaderOrder = true
    )
    {
        if (sheet is null) throw new ArgumentNullException(nameof(sheet));
        if (headerRowIndex < 0 || headerRowIndex >= sheet.RowCount)
            throw new ArgumentOutOfRangeException(nameof(headerRowIndex));

        headerNormalizer ??= static s => (s ?? string.Empty).Trim();
        cellString ??= DefaultCellToStringCollapseTrim;

        var comparer = caseInsensitiveHeaderMatch ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

        // 1) 讀表頭：HeaderText -> ColumnIndex
        var headerRow = sheet.Rows[headerRowIndex];
        var headerMap = new Dictionary<string, int>(comparer);
        var headerActual = new List<string>(headerRow.ColumnCount);
        for (int c = 0; c < headerRow.ColumnCount; c++)
        {
            var raw = cellString(headerRow.Cells[c]);
            var norm = headerNormalizer(raw);
            headerActual.Add(norm);
            if (!string.IsNullOrEmpty(norm) && !headerMap.ContainsKey(norm))
                headerMap[norm] = c;
        }

        // 2) 建 T 屬性對應（JsonPropertyName 或屬性名），並決定「期望順序」
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                             .Where(p => p.CanWrite)
                             .ToArray();

        var propMeta = props.Select(p => new
        {
            Property = p,
            JsonName = p.GetCustomAttribute<JsonPropertyAttribute>()?.PropertyName ?? p.Name,
            TargetType = p.PropertyType,
            Order = p.GetCustomAttribute<JsonPropertyAttribute>()?.Order
        }).ToArray();

        // 期望順序：先依 JsonPropertyOrder，再以 MetadataToken 近似宣告順序，最後以名稱穩定排序
        var orderedMeta = propMeta
            .OrderBy(m => m.Order ?? int.MaxValue)
            .ThenBy(m => m.Property.MetadataToken)
            .ThenBy(m => m.JsonName, StringComparer.Ordinal)
            .ToArray();

        var expectedNormalized = orderedMeta.Select(m => headerNormalizer(m.JsonName)).ToArray();

        // 缺少表頭 → 立即拋錯
        var missing = expectedNormalized.Where(h => !headerMap.ContainsKey(h)).ToArray();
        if (missing.Length > 0)
            throw new OdsHeaderMismatchException(expectedNormalized, missing, headerActual);

        // 欄位數量必須一致（以非空表頭為準）
        var actualNonEmpty = headerActual.Where(h => !string.IsNullOrEmpty(h)).ToArray();
        if (actualNonEmpty.Length != expectedNormalized.Length)
            throw new OdsHeaderCountMismatchException(expectedNormalized.Length, actualNonEmpty.Length, expectedNormalized, actualNonEmpty);

        // 檢查欄位順序（子序列相對順序必須一致；可容許中間有無關欄位）
        if (enforceHeaderOrder)
        {
            var expectedSet = new HashSet<string>(expectedNormalized, comparer);
            var actualSubsequence = headerActual.Where(h => expectedSet.Contains(h)).ToArray();
            if (!SequencesEqual(expectedNormalized, actualSubsequence, comparer))
            {
                throw new OdsHeaderOrderMismatchException(expectedNormalized, actualSubsequence, headerActual);
            }
        }

        // 3) 逐列轉換
        var results = new List<T>(Math.Max(0, sheet.RowCount - dataStartRowIndex));
        var errors = new List<OdsRowConversionException>();

        for (int r = Math.Max(dataStartRowIndex, headerRowIndex + 1); r < sheet.RowCount; r++)
        {
            var row = sheet.Rows[r];

            // 全空列判斷：以映射的欄位為準
            bool allEmpty = true;
            foreach (var m in orderedMeta)
            {
                var col = headerMap[headerNormalizer(m.JsonName)];
                var cell = col < row.ColumnCount ? row.Cells[col] : OdsCell.Empty;
                if (!IsEmpty(cell, cellString)) { allEmpty = false; break; }
            }
            if (allEmpty)
            {
                if (stopAtFirstAllEmptyRow) break;
                else continue;
            }

            var dict = new Dictionary<string, object?>(orderedMeta.Length, comparer);
            foreach (var m in orderedMeta)
            {
                var jsonName = m.JsonName;
                var normHeader = headerNormalizer(jsonName);
                var col = headerMap[normHeader];
                var headerName = headerRow.Cells[col].ToString();
                var cell = col < row.ColumnCount ? row.Cells[col] : OdsCell.Empty;

                try
                {
                    var value = ConvertCellToTarget(cell, m.TargetType, cellString);
                    // 以 JsonName（若有 [JsonProperty] 指定則使用）作為鍵，
                    // 讓 Newtonsoft.Json 依屬性對應名稱進行繫結，避免因覆寫名稱而無法賦值。
                    dict[jsonName] = value;
                }
                catch (System.Exception ex)
                {
                    errors.Add(new OdsRowConversionException(
                        rowIndex: r,
                        columnIndex: col,
                        jsonPropertyName: jsonName,
                        headerName: headerName,
                        targetType: m.TargetType,
                        cellType: cell.Type,
                        cellPreview: cell.ToString(),
                        inner: ex
                    ));
                }
            }

            if (errors.Count == 0 || errors[^1].RowIndex != r)
            {
                var json = JsonConvert.SerializeObject(dict);
                var obj = JsonConvert.DeserializeObject<T>(json, new JsonSerializerSettings
                {
                    // 允許字串與數值之間做寬鬆轉換
                    Culture = CultureInfo.InvariantCulture,
                    DateParseHandling = DateParseHandling.DateTimeOffset
                });
                results.Add(obj!);
            }
        }

        if (errors.Count > 0) throw new OdsAggregateConversionException(errors);
        return results;
    }

    // ===== Helper：順序比較、字串正規化、型別轉換、空值判斷 =====

    private static bool SequencesEqual(IEnumerable<string> a, IEnumerable<string> b, StringComparer cmp)
    {
        using var ea = a.GetEnumerator();
        using var eb = b.GetEnumerator();
        while (true)
        {
            var ma = ea.MoveNext();
            var mb = eb.MoveNext();
            if (ma != mb) return false; // 長度不同
            if (!ma) return true;       // 皆結束 → 相等
            if (cmp.Compare(ea.Current, eb.Current) != 0) return false;
        }
    }

    private static string DefaultCellToStringCollapseTrim(OdsCell c)
    {
        var s = c.Type == OdsValueType.String ? (c.Text ?? string.Empty) : c.ToString();
        if (string.IsNullOrEmpty(s)) return string.Empty;
        s = s.Replace("\r\n", "\n").Replace('\r', '\n');

        Span<char> buffer = stackalloc char[s.Length];
        int pos = 0; bool prevSpace = false;
        foreach (var ch in s)
        {
            bool spaceLike = ch == '\n' || ch == '\t' || ch == ' ';
            if (spaceLike)
            {
                if (!prevSpace)
                {
                    if (pos < buffer.Length) buffer[pos++] = ' ';
                    else return string.Join(' ', s.Split(new[] { '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries)).Trim();
                }
                prevSpace = true;
            }
            else
            {
                if (pos < buffer.Length) buffer[pos++] = ch;
                else return string.Join(' ', s.Split(new[] { '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries)).Trim();
                prevSpace = false;
            }
        }
        return new string(buffer[..pos]).Trim();
    }

    private static bool IsEmpty(in OdsCell c, Func<OdsCell, string> toText) => c.Type switch
    {
        OdsValueType.Empty => true,
        OdsValueType.String => string.IsNullOrWhiteSpace(toText(c)),
        OdsValueType.Float => !c.Number.HasValue,
        OdsValueType.Currency => !c.Number.HasValue,
        OdsValueType.Boolean => !c.Boolean.HasValue,
        OdsValueType.Date => !c.DateTime.HasValue,
        OdsValueType.Time => !c.Time.HasValue,
        _ => true
    };

    private static object? ConvertCellToTarget(in OdsCell c, Type targetType, Func<OdsCell, string> toText)
    {
        var u = Nullable.GetUnderlyingType(targetType);
        var isNullable = u is not null;
        var tt = u ?? targetType;

        if (IsEmpty(c, toText))
            return isNullable ? null : GetDefault(tt);

        if (tt == typeof(string)) return toText(c);

        // 數值系列
        if (tt == typeof(int) || tt == typeof(long) || tt == typeof(short) ||
            tt == typeof(double) || tt == typeof(float) || tt == typeof(decimal))
        {
            double? num = c.Number;
            if (!num.HasValue)
            {
                var s = toText(c);
                if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                    throw new FormatException($"無法將「{s}」解析為數字。");
                num = d;
            }
            checked
            {
                if (tt == typeof(int)) return (int)Math.Round(num.Value);
                if (tt == typeof(long)) return (long)Math.Round(num.Value);
                if (tt == typeof(short)) return (short)Math.Round(num.Value);
                if (tt == typeof(double)) return num.Value;
                if (tt == typeof(float)) return (float)num.Value;
                if (tt == typeof(decimal)) return (decimal)num.Value;
            }
        }

        // 布林
        if (tt == typeof(bool))
        {
            if (c.Type == OdsValueType.Boolean && c.Boolean.HasValue) return c.Boolean.Value;
            var s = toText(c);
            if (bool.TryParse(s, out var b)) return b;
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i)) return i != 0;
            throw new FormatException("此儲存格不是布林值。");
        }

        // 日期/時間
        if (tt == typeof(DateTimeOffset))
        {
            if (c.Type == OdsValueType.Date && c.DateTime.HasValue) return c.DateTime.Value;
            var s = toText(c);
            if (DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var dto))
                return dto;
            throw new FormatException($"無法解析日期時間「{s}」。");
        }
        if (tt == typeof(DateTime))
        {
            // 不轉換為本地或 UTC，維持原始牆鐘時間
            if (c.Type == OdsValueType.Date && c.DateTime.HasValue) return c.DateTime.Value.DateTime;
            var s = toText(c);
            // 優先用 DateTimeOffset 解析，最後取 .DateTime 以保留原始時間（Kind=Unspecified）
            if (DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var dto))
                return dto.DateTime;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                return dt;
            throw new FormatException($"無法解析日期時間「{s}」。");
        }
        if (tt == typeof(TimeSpan))
        {
            if (c.Type == OdsValueType.Time && c.Time.HasValue) return c.Time.Value;
            var s = toText(c);
            if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out var ts))
                return ts;
            throw new FormatException($"無法解析時間區段「{s}」。");
        }

        // Enum：名稱或數值
        if (tt.IsEnum)
        {
            var s = toText(c);
            if (Enum.TryParse(tt, s, ignoreCase: true, out var ev)) return ev!;
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ei))
                return Enum.ToObject(tt, ei);
            throw new FormatException($"無法解析列舉「{s}」為 {tt.Name}。");
        }

        // 其他型別：最後交給 Newtonsoft.Json 嘗試（以 JSON 字串進行寬鬆轉換）
        var json = JsonConvert.SerializeObject(c.ToString());
        return JsonConvert.DeserializeObject(json, tt);

        static object GetDefault(Type t) => t.IsValueType ? Activator.CreateInstance(t)! : null!;
    }

}
