using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using System.Threading.Tasks;
using NeoFileMagic.FileReader.Ods;
using Zaretto.ODS;


public sealed class TreatmentRow
{
    [JsonPropertyName("診療項目 代碼")]
    public string? Code { get; set; } = null;
    [JsonPropertyName("健保支付 點數")]
    public string? Points { get; set; } = null;
    [JsonPropertyName("生效起日")]
    public string? StartDate { get; set; } = null;
    [JsonPropertyName("生效迄日")]
    public string? EndDate { get; set; } = null;
    [JsonPropertyName("英文項目名稱")]
    public string? EnName { get; set; } = null;
    [JsonPropertyName("中文項目名稱")]
    public string? ZhName { get; set; } = null;
    [JsonPropertyName("備註")]
    public string? Note { get; set; } = null;
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

        var row0 = sheet.Rows[1];
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
        var row20 = sheet2.Rows[0];
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
        Console.WriteLine($"反序列化取得 {rows.Count()} 筆資料");
        // foreach (var row in rows)
        // {
        //     Console.WriteLine($"{row.Code}");
        //     Console.WriteLine($"{row.Points}");
        //     Console.WriteLine($"{row.StartDate}");
        //     Console.WriteLine($"{row.EndDate}");
        //     Console.WriteLine($"{row.EnName}");
        //     Console.WriteLine($"{row.ZhName}");
        //     Console.WriteLine($"{row.Note}");
        //     Console.WriteLine("---------------------");

        // }

        var odsReaderWriter = new ODSReaderWriter();
        var spreadsheetData = odsReaderWriter.ReadOdsFile("./sample.ods");
        DataTable table = spreadsheetData.Tables[0];
        System.Console.WriteLine("Sheet {0}", table.TableName);
        foreach (DataTable d in spreadsheetData.Tables)
        {
            System.Console.WriteLine("Sheet {0}", d.TableName);
            foreach (var row in d.AsEnumerable())
            {
                if (row.ItemArray.Any(ia => !string.IsNullOrEmpty(ia.ToString())))
                    System.Console.WriteLine("    {0}", string.Join(",", row.ItemArray.Select(xx => xx.ToString())));
            }
        }

    }
}
