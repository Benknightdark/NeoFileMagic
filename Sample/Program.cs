using System;
using System.Linq;
using System.Threading.Tasks;
using NeoFileMagic.FileReader.Ods;

class Program
{
    static async Task Main()
    {
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
    }
}
