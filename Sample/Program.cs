using System.Data;
using System.Text;
using Newtonsoft.Json;
using NeoFileMagic.FileReader.Ods;


public sealed class TreatmentRow
{
    [JsonProperty("診療項目 代碼")]
    public string Code { get; set; } = null!;
    [JsonProperty("健保支付 點數")]
    public string Points { get; set; } = null!;
    [JsonProperty("生效起日")]
    public DateTime StartDate { get; set; } 
    [JsonProperty("生效迄日")]
    public DateTime EndDate { get; set; } 
    [JsonProperty("英文項目名稱")]
    public string? EnName { get; set; } = null;
    [JsonProperty("中文項目名稱")]
    public string? ZhName { get; set; } = null;
    [JsonProperty("備註")]
    public string? Note { get; set; } = null;
}
class Program
{
    static async Task Main()
    {
        // 同步讀取本地 ODS
        var doc = NeoOds.Load("./sample.ods");
        var sheet = doc.Sheets[0];
        if (sheet.RowCount == 0)
        {
            Console.WriteLine("Sheet 為空");
            return;
        }

        var row0 = sheet.Rows[1];
        var values = row0.Cells.Select(c => c).ToArray();
        for (int i = 0; i < values.Length; i++)
        {
            Console.WriteLine($"[{i}] {values[i]}");
        }

        // 非同步下載並讀取 ODS
        var doc2 = await NeoOds.LoadFromUrlAsync(
            "https://info.nhi.gov.tw/api/iode0000s01/Dataset?rId=A21030000I-D20003-004",
            new OdsReaderOptions { LimitMode = OdsLimitMode.Truncate }
        );
        var sheet2 = doc2.Sheets[0];
        Console.WriteLine($"{sheet2.Name} Rows={sheet2.RowCount}");
        var row20 = sheet2.Rows[0];
        var values2 = row20.Cells.Select(c => c).ToArray();
        for (int i = 0; i < values2.Length; i++)
        {
            Console.WriteLine($"[{i}] {values2[i]}");
        }

        // 反序列化為物件集合
        var rows = NeoOds.DeserializeSheetOrThrow<TreatmentRow>(sheet2);
        Console.WriteLine($"反序列化取得 {rows.Count()} 筆資料");
        foreach (var row in rows)
        {
            Console.WriteLine($"{row.Code}");
            Console.WriteLine($"{row.Points}");
            Console.WriteLine($"{row.StartDate}");
            Console.WriteLine($"{row.EndDate}");
            Console.WriteLine($"{row.EnName}");
            Console.WriteLine($"{row.ZhName}");
            Console.WriteLine($"{row.Note}");
            Console.WriteLine("---------------------");
        }


        var jsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            StringEscapeHandling = StringEscapeHandling.Default
        };
        var outDir = "./out";
        Directory.CreateDirectory(outDir);
        var outPath = Path.Combine(outDir, "treatment_rows.json");
        var json = JsonConvert.SerializeObject(rows, jsonSettings);
        await File.WriteAllTextAsync(outPath, json, new UTF8Encoding(false));
        Console.WriteLine($"已輸出 JSON：{outPath}");

    }
}
