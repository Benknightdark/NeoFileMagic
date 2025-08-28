using System;
using System.Globalization;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using NeoFileMagic.FileReader.Ods;


public sealed class TreatmentRow
{
    [JsonPropertyName("診療項目 代碼")] public string? Code { get; set; } = null;
    [JsonPropertyName("健保支付 點數")] public string? Points { get; set; } = null;
    [JsonPropertyName("生效起日")] public string? StartDate { get; set; } = null;
    [JsonPropertyName("生效迄日")] public string? EndDate { get; set; } = null;
    [JsonPropertyName("英文項目名稱")] public string? EnName { get; set; } = null;
    [JsonPropertyName("中文項目名稱")] public string? ZhName { get; set; } = null;
    [JsonPropertyName("備註")] public string? Note { get; set; } = null;
}
class Program
{
    static string FormatIsoDateOrOriginal(string? s, string format = "yyyy-MM-dd")
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        if (DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture,
            DateTimeStyles.RoundtripKind, out var dto))
        {
            return dto.ToString(format, CultureInfo.InvariantCulture);
        }
        return s; // 非預期格式就如實輸出，避免例外
    }
    static async Task Main()
    {
        // 同步讀取本地 ODS
        var doc = Ods.Load("./sample.ods");
        var sheet = doc.Sheets[0];
        if (sheet.RowCount == 0)
        {
            Console.WriteLine("Sheet 為空");
            return;
        }

        var row0 = sheet.Rows[0];
        var values = row0.Cells.Select(c => c).ToArray();
        for (int i = 0; i < values.Length; i++)
        {
            Console.WriteLine($"[{i}] {values[i]}");
        }

        // 非同步下載並讀取 ODS
        var doc2 = await Ods.LoadFromUrlAsync(
            "https://info.nhi.gov.tw/api/iode0000s01/Dataset?rId=A21030000I-D20003-004",
            new OdsReaderOptions { LimitMode = OdsLimitMode.Truncate }
        );
        var sheet2 = doc2.Sheets[0];
        Console.WriteLine($"{sheet2.Name} Rows={sheet2.RowCount}");
        var row20 = sheet.Rows[0];
        var values2 = row20.Cells.Select(c => c).ToArray();
        for (int i = 0; i < values2.Length; i++)
        {
            Console.WriteLine($"[{i}] {values2[i]}");
        }

        // 反序列化為物件集合
        var rows = Ods.DeserializeSheetOrThrow<TreatmentRow>(
            sheet2,
            headerRowIndex: 0,
            dataStartRowIndex: 1,
            enforceHeaderOrder: true
        );
        Console.WriteLine($"反序列化取得 {rows.Count} 筆資料");
        foreach (var row in rows)
        {
            Console.WriteLine($"{row.Code} {row.ZhName} {row.Points} {FormatIsoDateOrOriginal(row.StartDate)}");
        }
    }
}
